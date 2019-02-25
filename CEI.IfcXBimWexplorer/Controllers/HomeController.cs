using CEI.IfcXBimWexplorer.Models;
using CEI.IfcXBimWexplorer.ViewModels;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Xbim.Common.Step21;
using Xbim.Ifc;
using Xbim.Ifc4.GeometricConstraintResource;
using Xbim.Ifc4.GeometricModelResource;
using Xbim.Ifc4.GeometryResource;
using Xbim.Ifc4.Interfaces;
using Xbim.Ifc4.Kernel;
using Xbim.Ifc4.MaterialResource;
using Xbim.Ifc4.MeasureResource;
using Xbim.Ifc4.PresentationAppearanceResource;
using Xbim.Ifc4.PresentationOrganizationResource;
using Xbim.Ifc4.ProductExtension;
using Xbim.Ifc4.ProfileResource;
using Xbim.Ifc4.PropertyResource;
using Xbim.Ifc4.RepresentationResource;
using Xbim.Ifc4.SharedBldgElements;
using Xbim.IO;
using Xbim.ModelGeometry.Scene;

namespace CEI.IfcXBimWexplorer.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult Viewer()
        {
            return View();
        }


        [HttpGet]
        public ActionResult Uploader()
        {
            //ColumnData columnData = new ColumnData();
            return View();
        }

        [HttpPost]
        public ActionResult Uploader(HttpPostedFileBase file, ColumnData columnData)
        {
            if (ModelState.IsValid && file != null && file.ContentLength > 0 && Path.GetExtension(file.FileName) == ".xlsx" && columnData.Height != 0 && columnData.Width != 0 && columnData.Length != 0)
            {
                //  Save Excel File
                var fileName = Path.GetFileNameWithoutExtension(file.FileName);
                var path = Path.Combine(Server.MapPath("~/data"), $"{fileName}.xlsx");
                file.SaveAs(path);

                //  Read From Excel File
                List<Point> pointsList = ReadFromTwoDimExcelFile(path);

                //  Create ifc-File Paths
                var ifcPath = Path.Combine(Server.MapPath("~/data"), $"{fileName}.ifc");
                var ifcXmlPath = Path.Combine(Server.MapPath("~/data"), $"{fileName}.ifcxml");

                //  Create IFc File
                CreateIFCFile("createdIFCFile", ifcPath, ifcXmlPath, columnData.Height, columnData.Width, columnData.Length, pointsList);

                //  Convert to WexBim
                ConvertToWexBIM(ifcPath);
                return RedirectToAction("Viewer");
            }

            else if (file != null && file.ContentLength > 0 && Path.GetExtension(file.FileName) == ".xlsx" && columnData.Height == 0 && columnData.Width == 0 && columnData.Length == 0)
            {
                //  Save Excel File
                var fileName = Path.GetFileNameWithoutExtension(file.FileName);
                var path = Path.Combine(Server.MapPath("~/data"), $"{fileName}.xlsx");
                file.SaveAs(path);

                //  Read From Excel File

                List<Frame> framesList = ReadFromThreeeDimExcelFile(path);

                //  Create ifc-File Paths
                var ifcPath = Path.Combine(Server.MapPath("~/data"), $"{fileName}.ifc");
                var ifcXmlPath = Path.Combine(Server.MapPath("~/data"), $"{fileName}.ifcxml");

                //  Create IFc File
                CreateIFCFile("createdIFCFile", ifcPath, ifcXmlPath, framesList);

                //  Convert to WexBim
                ConvertToWexBIM(ifcPath);
                return RedirectToAction("Viewer");
            }


            else if (file != null && file.ContentLength > 0 && Path.GetExtension(file.FileName) == ".ifc" && columnData.Height == 0 && columnData.Width == 0 && columnData.Length == 0)
            {
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/data"), fileName);
                file.SaveAs(path);
                ConvertToWexBIM(path);
                return RedirectToAction("Viewer");
            }


            else
            {
                return View();
            }


            //return RedirectToAction("Index", ViewBag);
        }

        #region Private Functions
        //----------------------------------  Step to XML   ----------------------------------//
        private static void ConvertToWexBIM(string filePath)
        {
            //  const string fileName = "../../my.ifc";
            using (var model = IfcStore.Open(filePath))
            {
                var context = new Xbim3DModelContext(model);
                context.CreateContext();


                // physical full path in drive
                var wexBimFullPath = Path.ChangeExtension(filePath, "wexBIM");

                var wexBimFileName = Path.GetFileName(wexBimFullPath);

                ConfigurationManager.AppSettings.Set("wexBIMFileName", wexBimFileName);
                ConfigurationManager.AppSettings.Set("wexBIMFullPath", "../data/" + wexBimFileName);

                using (var wexBiMfile = System.IO.File.Create(wexBimFullPath))
                {
                    using (var wexBimBinaryWriter = new BinaryWriter(wexBiMfile))
                    {
                        model.SaveAsWexBim(wexBimBinaryWriter);
                        wexBimBinaryWriter.Close();
                    }
                    wexBiMfile.Close();
                }

            }
        }

        //---------------------------------  Method Helpers for Creating IFC File  ----------------------------------//
        private static IfcStore CreateModelandAddProject(string projectName)
        {
            //  Setup Creadintials for ownership of the data in the new model
            var credentials = new XbimEditorCredentials()
            {
                ApplicationDevelopersName = "Everdawn-Studio",
                ApplicationFullName = projectName,
                ApplicationIdentifier = projectName + ".exe",
                ApplicationVersion = "1.0",
                EditorsFamilyName = "Track",
                EditorsGivenName = "CEI",
                EditorsOrganisationName = "ITI"
            };

            //  Create an Ifcstore in the memory
            var model = IfcStore.Create(credentials, IfcSchemaVersion.Ifc4, XbimStoreType.InMemoryModel);

            //  Begin Transaction as all change to a model are ACID
            using (var txn = model.BeginTransaction("Initialize Model"))
            {
                var project = model.Instances.New<IfcProject>(pr =>
                {
                    pr.Name = projectName;
                    pr.LongName = "longName__" + projectName;
                    pr.Phase = "Drafting";
                });

                #region Set the units of the project
                //project.Initialize(Xbim.Common.ProjectUnits.SIUnitsUK);

                //  Length
                project.UnitsInContext = model.Instances.New<IfcUnitAssignment>();
                var lengthUnit = model.Instances.New<IfcSIUnit>(siu =>
                {
                    siu.UnitType = IfcUnitEnum.LENGTHUNIT;  // mm
                    siu.Prefix = IfcSIPrefix.MILLI;
                    siu.Name = IfcSIUnitName.METRE;
                });
                project.UnitsInContext.Units.Add(lengthUnit);

                //  Area
                var areaUnit = model.Instances.New<IfcSIUnit>(siu =>
                {
                    siu.UnitType = IfcUnitEnum.AREAUNIT;
                    siu.Prefix = IfcSIPrefix.MICRO;
                    siu.Name = IfcSIUnitName.SQUARE_METRE;
                });
                project.UnitsInContext.Units.Add(areaUnit);

                //  Volume
                var volumeUnit = model.Instances.New<IfcSIUnit>(siu =>
                {
                    siu.UnitType = IfcUnitEnum.VOLUMEUNIT;
                    siu.Prefix = IfcSIPrefix.NANO;
                    siu.Name = IfcSIUnitName.CUBIC_METRE;
                });
                project.UnitsInContext.Units.Add(volumeUnit);

                //  Mass
                var massUnit = model.Instances.New<IfcSIUnit>(siu =>
                {
                    siu.UnitType = IfcUnitEnum.MASSUNIT;
                    siu.Prefix = IfcSIPrefix.GIGA;
                    siu.Name = IfcSIUnitName.GRAM;
                });
                project.UnitsInContext.Units.Add(massUnit);

                //  Mass-Density
                var massDensityUnit = model.Instances.New<IfcDerivedUnit>(du =>
                {
                    du.UnitType = IfcDerivedUnitEnum.MASSDENSITYUNIT;
                    du.Elements.Add(model.Instances.New<IfcDerivedUnitElement>(due =>
                    {
                        due.Exponent = 1;
                        due.Unit = massUnit;
                    }));
                    du.Elements.Add(model.Instances.New<IfcDerivedUnitElement>(due =>
                    {
                        due.Exponent = 1;
                        due.Unit = lengthUnit;
                    }));
                });
                project.UnitsInContext.Units.Add(massDensityUnit);

                //  sectionModulusUnit
                var sectionModulusUnit = model.Instances.New<IfcDerivedUnit>(du =>
                {
                    du.UnitType = IfcDerivedUnitEnum.SECTIONMODULUSUNIT;
                    du.Elements.Add(model.Instances.New<IfcDerivedUnitElement>(due =>
                    {
                        due.Exponent = 3;
                        due.Unit = lengthUnit;
                    }));
                });
                project.UnitsInContext.Units.Add(sectionModulusUnit);

                //  momentOfInertiaUnit
                var momentOfInertiaUnit = model.Instances.New<IfcDerivedUnit>(du =>
                {
                    du.UnitType = IfcDerivedUnitEnum.MOMENTOFINERTIAUNIT;
                    du.Elements.Add(model.Instances.New<IfcDerivedUnitElement>(due =>
                    {
                        due.Exponent = 4;
                        due.Unit = lengthUnit;
                    }));
                });
                project.UnitsInContext.Units.Add(momentOfInertiaUnit);

                //  planeAngleUnit
                var planeAngleUnit = model.Instances.New<IfcSIUnit>(siu =>
                {
                    siu.UnitType = IfcUnitEnum.PLANEANGLEUNIT;
                    siu.Name = IfcSIUnitName.RADIAN;
                });
                project.UnitsInContext.Units.Add(planeAngleUnit);

                //  timeUnit
                var timeUnit = model.Instances.New<IfcSIUnit>(siu =>
                {
                    siu.UnitType = IfcUnitEnum.TIMEUNIT;
                    siu.Name = IfcSIUnitName.SECOND;
                });
                project.UnitsInContext.Units.Add(timeUnit);

                //  forceUnit
                var forceUnit = model.Instances.New<IfcSIUnit>(siu =>
                {
                    siu.UnitType = IfcUnitEnum.FORCEUNIT;
                    siu.Prefix = IfcSIPrefix.KILO;
                    siu.Name = IfcSIUnitName.NEWTON;
                });
                project.UnitsInContext.Units.Add(forceUnit);

                //  torqueUnit
                var torqueUnit = model.Instances.New<IfcDerivedUnit>(du =>
                {
                    du.UnitType = IfcDerivedUnitEnum.TORQUEUNIT;
                    du.Elements.Add(model.Instances.New<IfcDerivedUnitElement>(due =>
                    {
                        due.Exponent = 1;
                        due.Unit = lengthUnit;
                    }));
                    du.Elements.Add(model.Instances.New<IfcDerivedUnitElement>(due =>
                    {
                        due.Exponent = 1;
                        due.Unit = forceUnit;
                    }));
                });
                project.UnitsInContext.Units.Add(torqueUnit);

                //  pressureUnit
                var pressureUnit = model.Instances.New<IfcSIUnit>(siu =>
                {
                    siu.UnitType = IfcUnitEnum.PRESSUREUNIT;
                    siu.Prefix = IfcSIPrefix.GIGA;
                    siu.Name = IfcSIUnitName.PASCAL;
                });
                project.UnitsInContext.Units.Add(pressureUnit);

                //  linearForceUnit
                var linearForceUnit = model.Instances.New<IfcDerivedUnit>(du =>
                {
                    du.UnitType = IfcDerivedUnitEnum.LINEARFORCEUNIT;
                    du.Elements.Add(model.Instances.New<IfcDerivedUnitElement>(due =>
                    {
                        due.Exponent = -1;
                        due.Unit = lengthUnit;
                    }));
                    du.Elements.Add(model.Instances.New<IfcDerivedUnitElement>(due =>
                    {
                        due.Exponent = 1;
                        due.Unit = forceUnit;
                    }));
                });
                project.UnitsInContext.Units.Add(linearForceUnit);

                //  planarForceUnit
                var planarForceUnit = model.Instances.New<IfcDerivedUnit>(du =>
                {
                    du.UnitType = IfcDerivedUnitEnum.PLANARFORCEUNIT;
                    du.Elements.Add(model.Instances.New<IfcDerivedUnitElement>(due =>
                    {
                        due.Exponent = -2;
                        due.Unit = lengthUnit;
                    }));
                    du.Elements.Add(model.Instances.New<IfcDerivedUnitElement>(due =>
                    {
                        due.Exponent = 1;
                        due.Unit = forceUnit;
                    }));
                });
                project.UnitsInContext.Units.Add(planarForceUnit);

                //  linearMomentUnit
                var linearMomentUnit = model.Instances.New<IfcDerivedUnit>(du =>
                {
                    du.UnitType = IfcDerivedUnitEnum.PLANARFORCEUNIT;
                    du.Elements.Add(model.Instances.New<IfcDerivedUnitElement>(due =>
                    {
                        due.Exponent = 1;
                        due.Unit = forceUnit;
                    }));
                    du.Elements.Add(model.Instances.New<IfcDerivedUnitElement>(due =>
                    {
                        due.Exponent = 1;
                        due.Unit = lengthUnit;
                    }));
                    du.Elements.Add(model.Instances.New<IfcDerivedUnitElement>(due =>
                    {
                        due.Exponent = -1;
                        due.Unit = lengthUnit;
                    }));
                });
                project.UnitsInContext.Units.Add(linearMomentUnit);

                //  shearModulusUnit
                var shearModulusUnit = model.Instances.New<IfcDerivedUnit>(du =>
                {
                    du.UnitType = IfcDerivedUnitEnum.SHEARMODULUSUNIT;
                    du.Elements.Add(model.Instances.New<IfcDerivedUnitElement>(due =>
                    {
                        due.Exponent = -2;
                        due.Unit = lengthUnit;
                    }));
                    du.Elements.Add(model.Instances.New<IfcDerivedUnitElement>(due =>
                    {
                        due.Exponent = 1;
                        due.Unit = forceUnit;
                    }));
                });
                project.UnitsInContext.Units.Add(shearModulusUnit);

                //  modulusOfElasticityUnit
                var modulusOfElasticityUnit = model.Instances.New<IfcDerivedUnit>(du =>
                {
                    du.UnitType = IfcDerivedUnitEnum.MODULUSOFELASTICITYUNIT;
                    du.Elements.Add(model.Instances.New<IfcDerivedUnitElement>(due =>
                    {
                        due.Exponent = -2;
                        due.Unit = lengthUnit;
                    }));
                    du.Elements.Add(model.Instances.New<IfcDerivedUnitElement>(due =>
                    {
                        due.Exponent = 1;
                        due.Unit = forceUnit;
                    }));
                });
                project.UnitsInContext.Units.Add(modulusOfElasticityUnit); ;

                //  thermodynamicTemperatureUnit
                var thermodynamicTemperatureUnit = model.Instances.New<IfcSIUnit>(siu =>
                {
                    siu.UnitType = IfcUnitEnum.THERMODYNAMICTEMPERATUREUNIT;
                    siu.Name = IfcSIUnitName.DEGREE_CELSIUS;
                });
                project.UnitsInContext.Units.Add(thermodynamicTemperatureUnit);

                //  thermalExpansionCoefficientUnit
                var thermalExpansionCoefficientUnit = model.Instances.New<IfcDerivedUnit>(du =>
                {
                    du.UnitType = IfcDerivedUnitEnum.MODULUSOFELASTICITYUNIT;
                    du.Elements.Add(model.Instances.New<IfcDerivedUnitElement>(due =>
                    {
                        due.Exponent = -1;
                        due.Unit = thermodynamicTemperatureUnit;
                    }));
                });
                project.UnitsInContext.Units.Add(thermalExpansionCoefficientUnit);
                #endregion

                txn.Commit();
            }
            return model;
        }

        private static IfcSite CreateSite(IfcStore model, string name)
        {
            using (var txn = model.BeginTransaction("Create Site"))
            {
                /*Creating a site instance*/
                var site = model.Instances.New<IfcSite>(s =>
                {
                    s.Name = name;
                    s.CompositionType = IfcElementCompositionEnum.ELEMENT;
                    s.RefLatitude = new IfcCompoundPlaneAngleMeasure(new List<long>() { 42, 21, 31, 181945 });
                    s.RefLongitude = new IfcCompoundPlaneAngleMeasure(new List<long>() { -71, -3, -24, -263305 });
                    s.RefElevation = 0;
                    s.ObjectPlacement = model.Instances.New<IfcLocalPlacement>(lp =>
                    {
                        lp.RelativePlacement = model.Instances.New<IfcAxis2Placement3D>(ap3d =>
                        {
                            //IfcCartesianPoint temp = model.Instances.OfType<IfcCartesianPoint>().Where(cp => Math.Abs(cp.X - 0)<0.0001 && cp.Y == 0 && cp.Z == 0).FirstOrDefault();

                            ap3d.Location = /*temp ??*/ model.Instances.New<IfcCartesianPoint>(cp => { cp.SetXYZ(0, 0, 0); });
                        });
                    });
                });


                var project = model.Instances.OfType<IfcProject>().FirstOrDefault();
                if (project != null) project.AddSite(site);

                txn.Commit();
                return site;
            }
        }

        private static IfcBuilding CreateBuilding(IfcStore model, string name)
        {
            using (var txn = model.BeginTransaction("Create Building"))
            {
                var site = model.Instances.OfType<IfcSite>().FirstOrDefault();
                if (site != null)
                {
                    /*Creating a building instance*/
                    var building = model.Instances.New<IfcBuilding>(bl =>
                    {
                        bl.Name = name;
                        bl.CompositionType = IfcElementCompositionEnum.ELEMENT;
                        bl.ObjectPlacement = model.Instances.New<IfcLocalPlacement>(lp =>
                        {
                            lp.RelativePlacement = model.Instances.New<IfcAxis2Placement3D>(ap3d => ap3d.Location = model.Instances.New<IfcCartesianPoint>(cp => { cp.SetXYZ(0, 0, 0); }));
                            lp.PlacementRelTo = site.ObjectPlacement;
                        });
                    });

                    site.AddBuilding(building);
                    txn.Commit();
                    return building;
                }
                else
                    return null;
            }
        }

        private static IfcBuildingStorey CreateStorey(IfcStore model, string name)
        {
            using (var txn = model.BeginTransaction("Create Storey"))
            {
                var building = model.Instances.OfType<IfcBuilding>().FirstOrDefault();
                if (building != null)
                {
                    /*Creating a storey instance*/
                    var storey = model.Instances.New<IfcBuildingStorey>(st =>
                    {
                        st.Name = name;
                        st.LongName = name;
                        st.CompositionType = IfcElementCompositionEnum.ELEMENT;
                        st.Elevation = 0;
                        st.ObjectPlacement = model.Instances.New<IfcLocalPlacement>(lp =>
                        {
                            //lp.PlacementRelTo = building.ObjectPlacement;
                            lp.RelativePlacement = model.Instances.New<IfcAxis2Placement3D>(ap3d =>
                                ap3d.Location = model.Instances.New<IfcCartesianPoint>(cp => { cp.SetXYZ(0, 0, 0); }));
                        });
                    });
                    building.AddToSpatialDecomposition(storey);
                    txn.Commit();
                    return storey;
                }
                else
                    return null;
            }
        }

        //---------------------------------  Create IFC Column  ----------------------------------//

        private static IfcColumn CreateIfcColumn(IfcStore model, IfcBuildingStorey storey, double length, double width, double xstart, double ystart, double zstart, double xend, double yend, double zend)
        {

            var columnName = $"StandardStrColumn_Rec_{width}_{length}";
            var columnObjectType = $"Rectangular Col {width} x {length}";

            var column = model.Instances.New<IfcColumn>(col =>
            {
                col.Name = columnName;
                col.ObjectType = columnObjectType;
                col.PredefinedType = IfcColumnTypeEnum.COLUMN;
                col.ObjectPlacement = model.Instances.New<IfcLocalPlacement>(lp =>
                {
                    lp.RelativePlacement = model.Instances.New<IfcAxis2Placement3D>(a2p3d =>
                    {
                        a2p3d.Location = model.Instances.New<IfcCartesianPoint>(loc => loc.SetXYZ(xstart, ystart, zstart));
                        //a2p3d.Axis = model.Instances.New<IfcDirection>(dir => dir.SetXYZ(0, 0, 1));
                        //a2p3d.RefDirection = model.Instances.New<IfcDirection>(dir => dir.SetXYZ(1, 0, 0));
                    });
                });
            });

            //-------------------------     Representation  -------------------------//
            #region ViualStyle
            //  Trying to find the CONCRETE visual style
            var surfaceStyleName = "Concrete, Cast-in-Place gray";
            var surfaceStyle = model.Instances.OfType<IfcSurfaceStyle>().Where(ss => ss.Name == surfaceStyleName).FirstOrDefault();
            if (surfaceStyle == null)
            {
                surfaceStyle = model.Instances.New<IfcSurfaceStyle>(ss =>
                {
                    ss.Name = surfaceStyleName;
                    ss.Side = IfcSurfaceSide.BOTH;
                    ss.Styles.Add(model.Instances.New<IfcSurfaceStyleRendering>(ssr =>
                    {
                        ssr.Transparency = 0;
                        ssr.ReflectanceMethod = IfcReflectanceMethodEnum.NOTDEFINED;
                        ssr.SurfaceColour = model.Instances.New<IfcColourRgb>(rgd =>
                        {
                            rgd.Red = 0.752941176470588;
                            rgd.Green = 0;
                            rgd.Blue = 0;
                        });
                        ssr.SpecularColour = new IfcNormalisedRatioMeasure(0.5);
                        ssr.SpecularHighlight = new IfcSpecularExponent(128);
                    }));
                });
            }

            //  Creating Styled-Item to be added to the extruded area solid
            var styledItem = model.Instances.New<IfcStyledItem>(si => si.Styles.Add(model.Instances.New<IfcPresentationStyleAssignment>(psa => psa.Styles.Add(surfaceStyle))));
            #endregion

            #region RepresentationContexts
            //  Try to find GeometricRepresentation-Context
            var geometricRepresentationContext = model.Instances.OfType<IfcProject>().FirstOrDefault().RepresentationContexts.Cast<IfcGeometricRepresentationContext>().Where(grc => grc.ContextIdentifier == null && grc.ContextType == "Model" && grc.CoordinateSpaceDimension == 3 && grc.Precision == 0.01).FirstOrDefault();

            if (geometricRepresentationContext == null)
            {
                geometricRepresentationContext = model.Instances.New<IfcGeometricRepresentationContext>(grc =>
                {
                    grc.ContextIdentifier = null;
                    grc.ContextType = "Model";
                    grc.CoordinateSpaceDimension = 3;
                    grc.Precision = 0.01;
                    grc.WorldCoordinateSystem = model.Instances.New<IfcAxis2Placement3D>(a2p =>
                    {
                        a2p.Location = model.Instances.New<IfcCartesianPoint>(cp => cp.SetXYZ(0, 0, 0));
                    });
                    grc.TrueNorth = model.Instances.New<IfcDirection>(dir => dir.SetXY(6.12303176911189E-17, 1));

                    model.Instances.OfType<IfcProject>().FirstOrDefault().RepresentationContexts.Add(grc);
                });
            }

            //  Try to find GeometricRepresentation-SubContext
            var geometricRepresentationSubContext = geometricRepresentationContext.HasSubContexts.Where(grsc => grsc.ContextIdentifier == "Body" && grsc.ContextType == "Model" && grsc.TargetView == IfcGeometricProjectionEnum.MODEL_VIEW).FirstOrDefault();
            if (geometricRepresentationSubContext == null)
            {
                geometricRepresentationSubContext = model.Instances.New<IfcGeometricRepresentationSubContext>(grsc =>
                {
                    grsc.ContextIdentifier = "Body";
                    grsc.ContextType = "Model";
                    grsc.TargetView = IfcGeometricProjectionEnum.MODEL_VIEW;

                    grsc.ParentContext = geometricRepresentationContext;    //  Adding to parent Context
                });
            }
            #endregion

            #region Completing Representation Steps
            //  PresentationLayerAssignment CAD presentation 
            //  Add the created shape representations to it
            var presentationLayerAssignment = model.Instances.OfType<IfcPresentationLayerAssignment>().Where(pla => pla.Name == "S-COLS").FirstOrDefault();
            if (presentationLayerAssignment == null)
                presentationLayerAssignment = model.Instances.New<IfcPresentationLayerAssignment>(pla => pla.Name = "S-COLS");

            var ProfileName = $"RectProfile{width}x{length}";

            //  Creating the Representation Map
            var representationMap = model.Instances.New<IfcRepresentationMap>(rm =>
            {
                rm.MappingOrigin = model.Instances.New<IfcAxis2Placement3D>(a2p3d => a2p3d.Location = model.Instances.New<IfcCartesianPoint>(cp => cp.SetXYZ(0, 0, 0)));
                rm.MappedRepresentation = model.Instances.New<IfcShapeRepresentation>(msr =>
                {
                    presentationLayerAssignment.AssignedItems.Add(msr);
                    msr.RepresentationIdentifier = "Body";
                    msr.RepresentationType = "SweptSolid";
                    msr.ContextOfItems = geometricRepresentationSubContext;
                    msr.Items.Add(model.Instances.New<IfcExtrudedAreaSolid>(eas =>
                    {
                        if (xstart == xend && ystart == yend)
                        {
                            eas.Depth = Math.Abs(zend - zstart);
                        }
                        else if (xstart == xend && zstart == zend)
                        {
                            eas.Depth = Math.Abs(yend - ystart);
                        }
                        else if (ystart == yend && zstart == zend)
                        {
                            eas.Depth = Math.Abs(xend - xstart);
                        }

                        eas.SweptArea = model.Instances.New<IfcRectangleProfileDef>(rpd =>
                        {
                            rpd.ProfileType = IfcProfileTypeEnum.AREA;
                            rpd.ProfileName = ProfileName;
                            rpd.YDim = width;
                            rpd.XDim = length;

                            if ((xstart == xend && ystart == yend) || (xstart == xend && zstart == zend))
                            {
                                rpd.Position = model.Instances.New<IfcAxis2Placement2D>(a2p2 =>
                                {
                                    a2p2.Location = model.Instances.New<IfcCartesianPoint>(mcp => mcp.SetXY(0, 0));
                                    a2p2.RefDirection = model.Instances.New<IfcDirection>(md => md.SetXY(1, 0));

                                });
                            }

                            else if (ystart == yend && zstart == zend)
                            {
                                rpd.Position = model.Instances.New<IfcAxis2Placement2D>(a2p2 =>
                                {
                                    a2p2.Location = model.Instances.New<IfcCartesianPoint>(mcp => mcp.SetXY(0, 0));
                                    a2p2.RefDirection = model.Instances.New<IfcDirection>(md => md.SetXY(0, 1));
                                });
                            }

                        });

                        eas.Position = model.Instances.New<IfcAxis2Placement3D>(a2pl3d =>
                        {
                            if (xstart == xend && ystart == yend)
                            {
                                a2pl3d.Location = model.Instances.New<IfcCartesianPoint>(cp => { cp.SetXYZ(0, 0, 0); });
                                a2pl3d.Axis = model.Instances.New<IfcDirection>(a2ad => a2ad.SetXYZ(0, 0, 1));
                                a2pl3d.RefDirection = model.Instances.New<IfcDirection>(a2rd => a2rd.SetXYZ(0, -1, 0));
                            }
                            else if (xstart == xend && zstart == zend)
                            {
                                a2pl3d.Location = model.Instances.New<IfcCartesianPoint>(cp => { cp.SetXYZ(0, 0, 0); });
                                a2pl3d.Axis = model.Instances.New<IfcDirection>(a2ad => a2ad.SetXYZ(1, 0, 0));
                                a2pl3d.RefDirection = model.Instances.New<IfcDirection>(a2rd => a2rd.SetXYZ(0, 1, 0));
                            }
                            else if (ystart == yend && zstart == zend)
                            {
                                a2pl3d.Location = model.Instances.New<IfcCartesianPoint>(cp => { cp.SetXYZ(0, 0, 0); });
                                a2pl3d.Axis = model.Instances.New<IfcDirection>(a2ad => a2ad.SetXYZ(0, 1, 0));
                                a2pl3d.RefDirection = model.Instances.New<IfcDirection>(a2rd => a2rd.SetXYZ(0, 0, 1));
                            }

                        });
                        eas.ExtrudedDirection = eas.Position.Axis;

                        styledItem.Item = eas; //Adding the visual style to the solid 
                    }));
                });
            });

            //  Creating Column Representation and adding the representation map to it
            column.Representation = model.Instances.New<IfcProductDefinitionShape>(pds =>
            {
                pds.Representations.Add(model.Instances.New<IfcShapeRepresentation>(sr =>
                {
                    presentationLayerAssignment.AssignedItems.Add(sr);
                    sr.RepresentationIdentifier = "Body";
                    sr.RepresentationType = "MappedRepresentation";
                    sr.ContextOfItems = geometricRepresentationSubContext;
                    sr.Items.Add(model.Instances.New<IfcMappedItem>(mi =>
                    {
                        mi.MappingSource = representationMap;
                        mi.MappingTarget = model.Instances.New<IfcCartesianTransformationOperator3D>(cto3 =>
                        {
                            cto3.Scale = 1;
                            cto3.LocalOrigin = model.Instances.New<IfcCartesianPoint>(cp => cp.SetXYZ(0, 0, 0));
                        });
                    }));
                }));
            });

            //  Creating a Relation between the COLUMN and its COLUMN-Type using the Representation-Map
            var columnTypeRelation = model.Instances.New<IfcRelDefinesByType>(rdbt =>
            {
                rdbt.RelatingType = model.Instances.New<IfcColumnType>(ct =>
                {
                    ct.Name = $"RectColumn{width}x{length}";
                    ct.ElementType = $"RectColumn{width}x{length}";
                    ct.PredefinedType = IfcColumnTypeEnum.COLUMN;
                    ct.RepresentationMaps.Add(representationMap);
                });
            });
            columnTypeRelation.RelatedObjects.Add(column);
            #endregion

            // Adding some properties to the column
            var relDefineByProps = model.Instances.New<IfcRelDefinesByProperties>(rdbp =>
            {
                rdbp.RelatedObjects.Add(column);
                rdbp.RelatingPropertyDefinition = model.Instances.New<IfcPropertySet>(ps =>
                {
                    ps.Name = "Column Property Set";
                    ps.HasProperties.Add(model.Instances.New<IfcPropertySingleValue>(psv =>
                    {
                        psv.Name = "LoadBearing";
                        psv.NominalValue = new IfcBoolean(true);
                    }));
                    ps.HasProperties.Add(model.Instances.New<IfcPropertySingleValue>(psv =>
                    {
                        psv.Name = "Reference";
                        psv.NominalValue = new IfcIdentifier(ProfileName);
                    }));
                });
            });

            //  Adding Material 
            #region Query for units from Project
            //----------------------------------------------------------
            //Ifc SI Units
            //----------------------------------------------------------

            var lengthUnit = model.Instances.OfType<IfcSIUnit>().Where(siu =>
                siu.UnitType == IfcUnitEnum.LENGTHUNIT
            ).FirstOrDefault();

            var areaUnit = model.Instances.OfType<IfcSIUnit>().Where(siu =>
                siu.UnitType == IfcUnitEnum.AREAUNIT
            ).FirstOrDefault();

            var volumeUnit = model.Instances.OfType<IfcSIUnit>().Where(siu =>
                siu.UnitType == IfcUnitEnum.VOLUMEUNIT
            ).FirstOrDefault();

            var sectionModulusUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.SECTIONMODULUSUNIT
            ).FirstOrDefault();

            var momentOfInertiaUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.MOMENTOFINERTIAUNIT
            ).FirstOrDefault();

            var planeAngleUnit = model.Instances.OfType<IfcSIUnit>().Where(siu =>
                siu.UnitType == IfcUnitEnum.PLANEANGLEUNIT
            ).FirstOrDefault();

            var timeUnit = model.Instances.OfType<IfcSIUnit>().Where(siu =>
                siu.UnitType == IfcUnitEnum.TIMEUNIT
            ).FirstOrDefault();

            var massUnit = model.Instances.OfType<IfcSIUnit>().Where(siu =>
                siu.UnitType == IfcUnitEnum.MASSUNIT
            ).FirstOrDefault();

            var massDensityUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.MASSDENSITYUNIT
            ).FirstOrDefault();

            var forceUnit = model.Instances.OfType<IfcSIUnit>().Where(siu =>
                siu.UnitType == IfcUnitEnum.FORCEUNIT
            ).FirstOrDefault();

            var torqueUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.TORQUEUNIT
            ).FirstOrDefault();

            var pressureUnit = model.Instances.OfType<IfcSIUnit>().Where(siu =>
                siu.UnitType == IfcUnitEnum.PRESSUREUNIT
            ).FirstOrDefault();

            var linearForceUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.LINEARFORCEUNIT
            ).FirstOrDefault();

            var planarForceUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.PLANARFORCEUNIT
            ).FirstOrDefault();

            var linearMomentUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.PLANARFORCEUNIT
            ).FirstOrDefault();

            var shearModulusUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.SHEARMODULUSUNIT
            ).FirstOrDefault();

            var modulusOfElasticityUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.MODULUSOFELASTICITYUNIT
            ).FirstOrDefault();

            var thermodynamicTemperatureUnit = model.Instances.OfType<IfcSIUnit>().Where(siu =>
                siu.UnitType == IfcUnitEnum.THERMODYNAMICTEMPERATUREUNIT
            ).FirstOrDefault();

            var thermalExpansionCoefficientUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.MODULUSOFELASTICITYUNIT
            ).FirstOrDefault();

            //-------------------------------------------------------------------
            #endregion

            #region CreatingMaterial using selected units
            var materialName = "CompleteMaterial";
            var materialProfileSetUsage = model.Instances.New<IfcMaterialProfileSetUsage>(mpsu =>
            {
                mpsu.ForProfileSet = model.Instances.New<IfcMaterialProfileSet>(mps =>
                {
                    mps.MaterialProfiles.Add(model.Instances.New<IfcMaterialProfile>(mp =>
                    {
                        mp.Material = model.Instances.New<IfcMaterial>(mat => mat.Name = materialName);
                    }));
                });
            });

            //  Creating Material Representation for cad styling
            var materialRepresentation = model.Instances.New<IfcMaterialDefinitionRepresentation>(mdr =>
            {
                mdr.Representations.Add(model.Instances.New<IfcStyledRepresentation>(sr =>
                {
                    sr.RepresentationIdentifier = "Style";
                    sr.RepresentationType = "Material";
                    sr.ContextOfItems = geometricRepresentationContext;
                    sr.Items.Add(model.Instances.New<IfcStyledItem>(si =>
                        si.Styles.Add(model.Instances.New<IfcPresentationStyleAssignment>(psa => psa.Styles.Add(surfaceStyle)))) // Adding the Surface Style to the Material 
                    );
                }));
                mdr.RepresentedMaterial = materialProfileSetUsage.ForProfileSet.MaterialProfiles.First.Material;
            });

            //  IfcMaterialProperties and Adding the created ifc material to it
            var materialProperties = model.Instances.New<IfcMaterialProperties>(mp =>
            {
                mp.Name = materialName; ;
                mp.Material = materialProfileSetUsage.ForProfileSet.MaterialProfiles.First.Material;

                mp.Properties.Add(model.Instances.New<IfcPropertySingleValue>(prop =>
                {
                    prop.Name = "CompressiveStrength";
                    prop.NominalValue = new IfcPressureMeasure(0.037579032);
                    prop.Unit = pressureUnit;
                }));
                mp.Properties.Add(model.Instances.New<IfcPropertySingleValue>(prop =>
                {
                    prop.Name = "MassDensity";
                    prop.NominalValue = new IfcMassDensityMeasure(3.4027e-12);
                    prop.Unit = massDensityUnit;
                }));
                mp.Properties.Add(model.Instances.New<IfcPropertySingleValue>(prop =>
                {
                    prop.Name = "PoissonRatio";
                    prop.NominalValue = new IfcRatioMeasure(0.3);
                }));
                mp.Properties.Add(model.Instances.New<IfcPropertySingleValue>(prop =>
                {
                    prop.Name = "ThermalExpansionCoefficient";
                    prop.NominalValue = new IfcThermalExpansionCoefficientMeasure(10.8999999e-6);
                    prop.Unit = thermalExpansionCoefficientUnit;
                }));
                mp.Properties.Add(model.Instances.New<IfcPropertySingleValue>(prop =>
                {
                    prop.Name = "YoungModulus";
                    prop.NominalValue = new IfcModulusOfElasticityMeasure(24.855578);
                    prop.Unit = modulusOfElasticityUnit;
                }));

            });
            #endregion

            #region Assigning The material to the Column
            var relAssociatesMaterial = model.Instances.New<IfcRelAssociatesMaterial>();

            relAssociatesMaterial.RelatingMaterial = materialProfileSetUsage.ForProfileSet.MaterialProfiles.First.Material;
            relAssociatesMaterial.RelatedObjects.Add(columnTypeRelation.RelatingType);
            relAssociatesMaterial.RelatedObjects.Add(column);
            #endregion


            //  Add to the storey
            //var storey = model.Instances.OfType<IfcBuildingStorey>().FirstOrDefault();
            //if (storey != null)
            storey.AddElement(column);

            //txn.Commit();
            return column;
        }

        private static IfcColumn CreateIfcColumn(IfcStore model, IfcBuildingStorey storey, double length, double width, double height, double x, double y, double z)
        {
            //using (var txn = model.BeginTransaction())
            //{
            //}

            var columnName = $"StandardStrColumn_Rec_{width}_{length}";
            var columnObjectType = $"Rectangular Col {width} x {length}";

            var column = model.Instances.New<IfcColumn>(col =>
            {
                col.Name = columnName;
                col.ObjectType = columnObjectType;
                col.PredefinedType = IfcColumnTypeEnum.COLUMN;
                col.ObjectPlacement = model.Instances.New<IfcLocalPlacement>(lp =>
                {
                    lp.RelativePlacement = model.Instances.New<IfcAxis2Placement3D>(a2p3d =>
                    {
                        a2p3d.Location = model.Instances.New<IfcCartesianPoint>(loc => loc.SetXYZ(x, y, z));
                        //a2p3d.Axis = model.Instances.New<IfcDirection>(dir => dir.SetXYZ(0, 0, 1));
                        //a2p3d.RefDirection = model.Instances.New<IfcDirection>(dir => dir.SetXYZ(1, 0, 0));
                    });
                });
            });

            //-------------------------     Representation  -------------------------//
            #region ViualStyle
            //  Trying to find the CONCRETE visual style
            var surfaceStyleName = "Concrete, Cast-in-Place gray";
            var surfaceStyle = model.Instances.OfType<IfcSurfaceStyle>().Where(ss => ss.Name == surfaceStyleName).FirstOrDefault();
            if (surfaceStyle == null)
            {
                surfaceStyle = model.Instances.New<IfcSurfaceStyle>(ss =>
                {
                    ss.Name = surfaceStyleName;
                    ss.Side = IfcSurfaceSide.BOTH;
                    ss.Styles.Add(model.Instances.New<IfcSurfaceStyleRendering>(ssr =>
                    {
                        ssr.Transparency = 0;
                        ssr.ReflectanceMethod = IfcReflectanceMethodEnum.NOTDEFINED;
                        ssr.SurfaceColour = model.Instances.New<IfcColourRgb>(rgd =>
                        {
                            rgd.Red = 0.752941176470588;
                            rgd.Green = 0;
                            rgd.Blue = 0;
                        });
                        ssr.SpecularColour = new IfcNormalisedRatioMeasure(0.5);
                        ssr.SpecularHighlight = new IfcSpecularExponent(128);
                    }));
                });
            }

            //  Creating Styled-Item to be added to the extruded area solid
            var styledItem = model.Instances.New<IfcStyledItem>(si => si.Styles.Add(model.Instances.New<IfcPresentationStyleAssignment>(psa => psa.Styles.Add(surfaceStyle))));
            #endregion

            #region RepresentationContexts
            //  Try to find GeometricRepresentation-Context
            var geometricRepresentationContext = model.Instances.OfType<IfcProject>().FirstOrDefault().RepresentationContexts.Cast<IfcGeometricRepresentationContext>().Where(grc => grc.ContextIdentifier == null && grc.ContextType == "Model" && grc.CoordinateSpaceDimension == 3 && grc.Precision == 0.01).FirstOrDefault();

            if (geometricRepresentationContext == null)
            {
                geometricRepresentationContext = model.Instances.New<IfcGeometricRepresentationContext>(grc =>
                {
                    grc.ContextIdentifier = null;
                    grc.ContextType = "Model";
                    grc.CoordinateSpaceDimension = 3;
                    grc.Precision = 0.01;
                    grc.WorldCoordinateSystem = model.Instances.New<IfcAxis2Placement3D>(a2p =>
                    {
                        a2p.Location = model.Instances.New<IfcCartesianPoint>(cp => cp.SetXYZ(0, 0, 0));
                    });
                    grc.TrueNorth = model.Instances.New<IfcDirection>(dir => dir.SetXY(6.12303176911189E-17, 1));

                    model.Instances.OfType<IfcProject>().FirstOrDefault().RepresentationContexts.Add(grc);
                });
            }

            //  Try to find GeometricRepresentation-SubContext
            var geometricRepresentationSubContext = geometricRepresentationContext.HasSubContexts.Where(grsc => grsc.ContextIdentifier == "Body" && grsc.ContextType == "Model" && grsc.TargetView == IfcGeometricProjectionEnum.MODEL_VIEW).FirstOrDefault();
            if (geometricRepresentationSubContext == null)
            {
                geometricRepresentationSubContext = model.Instances.New<IfcGeometricRepresentationSubContext>(grsc =>
                {
                    grsc.ContextIdentifier = "Body";
                    grsc.ContextType = "Model";
                    grsc.TargetView = IfcGeometricProjectionEnum.MODEL_VIEW;

                    grsc.ParentContext = geometricRepresentationContext;    //  Adding to parent Context
                });
            }
            #endregion

            #region Completing Representation Steps
            //  PresentationLayerAssignment CAD presentation 
            //  Add the created shape representations to it
            var presentationLayerAssignment = model.Instances.OfType<IfcPresentationLayerAssignment>().Where(pla => pla.Name == "S-COLS").FirstOrDefault();
            if (presentationLayerAssignment == null)
                presentationLayerAssignment = model.Instances.New<IfcPresentationLayerAssignment>(pla => pla.Name = "S-COLS");

            var ProfileName = $"RectProfile{width}x{length}";

            //  Creating the Representation Map
            var representationMap = model.Instances.New<IfcRepresentationMap>(rm =>
            {
                rm.MappingOrigin = model.Instances.New<IfcAxis2Placement3D>(a2p3d => a2p3d.Location = model.Instances.New<IfcCartesianPoint>(cp => cp.SetXYZ(0, 0, 0)));
                rm.MappedRepresentation = model.Instances.New<IfcShapeRepresentation>(msr =>
                {
                    presentationLayerAssignment.AssignedItems.Add(msr);
                    msr.RepresentationIdentifier = "Body";
                    msr.RepresentationType = "SweptSolid";
                    msr.ContextOfItems = geometricRepresentationSubContext;
                    msr.Items.Add(model.Instances.New<IfcExtrudedAreaSolid>(eas =>
                    {
                        eas.Depth = height;
                        eas.SweptArea = model.Instances.New<IfcRectangleProfileDef>(rpd =>
                        {
                            rpd.ProfileType = IfcProfileTypeEnum.AREA;
                            rpd.ProfileName = ProfileName;
                            rpd.YDim = length;
                            rpd.XDim = width;
                            rpd.Position = model.Instances.New<IfcAxis2Placement2D>(a2p2 =>
                            {
                                a2p2.Location = model.Instances.New<IfcCartesianPoint>(mcp => mcp.SetXY(0, 0));
                                a2p2.RefDirection = model.Instances.New<IfcDirection>(md => md.SetXY(1, 0));
                            });
                        });
                        eas.Position = model.Instances.New<IfcAxis2Placement3D>(a2pl3d =>
                        {
                            a2pl3d.Location = model.Instances.New<IfcCartesianPoint>(cp => { cp.SetXYZ(0, 0, 0); });
                            a2pl3d.Axis = model.Instances.New<IfcDirection>(a2ad => a2ad.SetXYZ(0, 0, 1));
                            a2pl3d.RefDirection = model.Instances.New<IfcDirection>(a2rd => a2rd.SetXYZ(0, -1, 0));
                        });
                        eas.ExtrudedDirection = eas.Position.Axis;

                        styledItem.Item = eas; //Adding the visual style to the solid 
                    }));
                });
            });

            //  Creating Column Representation and adding the representation map to it
            column.Representation = model.Instances.New<IfcProductDefinitionShape>(pds =>
            {
                pds.Representations.Add(model.Instances.New<IfcShapeRepresentation>(sr =>
                {
                    presentationLayerAssignment.AssignedItems.Add(sr);
                    sr.RepresentationIdentifier = "Body";
                    sr.RepresentationType = "MappedRepresentation";
                    sr.ContextOfItems = geometricRepresentationSubContext;
                    sr.Items.Add(model.Instances.New<IfcMappedItem>(mi =>
                    {
                        mi.MappingSource = representationMap;
                        mi.MappingTarget = model.Instances.New<IfcCartesianTransformationOperator3D>(cto3 =>
                        {
                            cto3.Scale = 1;
                            cto3.LocalOrigin = model.Instances.New<IfcCartesianPoint>(cp => cp.SetXYZ(0, 0, 0));
                        });
                    }));
                }));
            });

            //  Creating a Relation between the COLUMN and its COLUMN-Type using the Representation-Map
            var columnTypeRelation = model.Instances.New<IfcRelDefinesByType>(rdbt =>
            {
                rdbt.RelatingType = model.Instances.New<IfcColumnType>(ct =>
                {
                    ct.Name = $"RectColumn{width}x{length}";
                    ct.ElementType = $"RectColumn{width}x{length}";
                    ct.PredefinedType = IfcColumnTypeEnum.COLUMN;
                    ct.RepresentationMaps.Add(representationMap);
                });
            });
            columnTypeRelation.RelatedObjects.Add(column);
            #endregion

            // Adding some properties to the column
            var relDefineByProps = model.Instances.New<IfcRelDefinesByProperties>(rdbp =>
            {
                rdbp.RelatedObjects.Add(column);
                rdbp.RelatingPropertyDefinition = model.Instances.New<IfcPropertySet>(ps =>
                {
                    ps.Name = "Column Property Set";
                    ps.HasProperties.Add(model.Instances.New<IfcPropertySingleValue>(psv =>
                    {
                        psv.Name = "LoadBearing";
                        psv.NominalValue = new IfcBoolean(true);
                    }));
                    ps.HasProperties.Add(model.Instances.New<IfcPropertySingleValue>(psv =>
                    {
                        psv.Name = "Reference";
                        psv.NominalValue = new IfcIdentifier(ProfileName);
                    }));
                });
            });

            //  Adding Material 
            #region Query for units from Project
            //----------------------------------------------------------
            //Ifc SI Units
            //----------------------------------------------------------

            var lengthUnit = model.Instances.OfType<IfcSIUnit>().Where(siu =>
                siu.UnitType == IfcUnitEnum.LENGTHUNIT
            ).FirstOrDefault();

            var areaUnit = model.Instances.OfType<IfcSIUnit>().Where(siu =>
                siu.UnitType == IfcUnitEnum.AREAUNIT
            ).FirstOrDefault();

            var volumeUnit = model.Instances.OfType<IfcSIUnit>().Where(siu =>
                siu.UnitType == IfcUnitEnum.VOLUMEUNIT
            ).FirstOrDefault();

            var sectionModulusUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.SECTIONMODULUSUNIT
            ).FirstOrDefault();

            var momentOfInertiaUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.MOMENTOFINERTIAUNIT
            ).FirstOrDefault();

            var planeAngleUnit = model.Instances.OfType<IfcSIUnit>().Where(siu =>
                siu.UnitType == IfcUnitEnum.PLANEANGLEUNIT
            ).FirstOrDefault();

            var timeUnit = model.Instances.OfType<IfcSIUnit>().Where(siu =>
                siu.UnitType == IfcUnitEnum.TIMEUNIT
            ).FirstOrDefault();

            var massUnit = model.Instances.OfType<IfcSIUnit>().Where(siu =>
                siu.UnitType == IfcUnitEnum.MASSUNIT
            ).FirstOrDefault();

            var massDensityUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.MASSDENSITYUNIT
            ).FirstOrDefault();

            var forceUnit = model.Instances.OfType<IfcSIUnit>().Where(siu =>
                siu.UnitType == IfcUnitEnum.FORCEUNIT
            ).FirstOrDefault();

            var torqueUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.TORQUEUNIT
            ).FirstOrDefault();

            var pressureUnit = model.Instances.OfType<IfcSIUnit>().Where(siu =>
                siu.UnitType == IfcUnitEnum.PRESSUREUNIT
            ).FirstOrDefault();

            var linearForceUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.LINEARFORCEUNIT
            ).FirstOrDefault();

            var planarForceUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.PLANARFORCEUNIT
            ).FirstOrDefault();

            var linearMomentUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.PLANARFORCEUNIT
            ).FirstOrDefault();

            var shearModulusUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.SHEARMODULUSUNIT
            ).FirstOrDefault();

            var modulusOfElasticityUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.MODULUSOFELASTICITYUNIT
            ).FirstOrDefault();

            var thermodynamicTemperatureUnit = model.Instances.OfType<IfcSIUnit>().Where(siu =>
                siu.UnitType == IfcUnitEnum.THERMODYNAMICTEMPERATUREUNIT
            ).FirstOrDefault();

            var thermalExpansionCoefficientUnit = model.Instances.OfType<IfcDerivedUnit>().Where(du =>
                du.UnitType == IfcDerivedUnitEnum.MODULUSOFELASTICITYUNIT
            ).FirstOrDefault();

            //-------------------------------------------------------------------
            #endregion

            #region CreatingMaterial using selected units
            var materialName = "CompleteMaterial";
            var materialProfileSetUsage = model.Instances.New<IfcMaterialProfileSetUsage>(mpsu =>
            {
                mpsu.ForProfileSet = model.Instances.New<IfcMaterialProfileSet>(mps =>
                {
                    mps.MaterialProfiles.Add(model.Instances.New<IfcMaterialProfile>(mp =>
                    {
                        mp.Material = model.Instances.New<IfcMaterial>(mat => mat.Name = materialName);
                    }));
                });
            });

            //  Creating Material Representation for cad styling
            var materialRepresentation = model.Instances.New<IfcMaterialDefinitionRepresentation>(mdr =>
            {
                mdr.Representations.Add(model.Instances.New<IfcStyledRepresentation>(sr =>
                {
                    sr.RepresentationIdentifier = "Style";
                    sr.RepresentationType = "Material";
                    sr.ContextOfItems = geometricRepresentationContext;
                    sr.Items.Add(model.Instances.New<IfcStyledItem>(si =>
                        si.Styles.Add(model.Instances.New<IfcPresentationStyleAssignment>(psa => psa.Styles.Add(surfaceStyle)))) // Adding the Surface Style to the Material 
                    );
                }));
                mdr.RepresentedMaterial = materialProfileSetUsage.ForProfileSet.MaterialProfiles.First.Material;
            });

            //  IfcMaterialProperties and Adding the created ifc material to it
            var materialProperties = model.Instances.New<IfcMaterialProperties>(mp =>
            {
                mp.Name = materialName; ;
                mp.Material = materialProfileSetUsage.ForProfileSet.MaterialProfiles.First.Material;

                mp.Properties.Add(model.Instances.New<IfcPropertySingleValue>(prop =>
                {
                    prop.Name = "CompressiveStrength";
                    prop.NominalValue = new IfcPressureMeasure(0.037579032);
                    prop.Unit = pressureUnit;
                }));
                mp.Properties.Add(model.Instances.New<IfcPropertySingleValue>(prop =>
                {
                    prop.Name = "MassDensity";
                    prop.NominalValue = new IfcMassDensityMeasure(3.4027e-12);
                    prop.Unit = massDensityUnit;
                }));
                mp.Properties.Add(model.Instances.New<IfcPropertySingleValue>(prop =>
                {
                    prop.Name = "PoissonRatio";
                    prop.NominalValue = new IfcRatioMeasure(0.3);
                }));
                mp.Properties.Add(model.Instances.New<IfcPropertySingleValue>(prop =>
                {
                    prop.Name = "ThermalExpansionCoefficient";
                    prop.NominalValue = new IfcThermalExpansionCoefficientMeasure(10.8999999e-6);
                    prop.Unit = thermalExpansionCoefficientUnit;
                }));
                mp.Properties.Add(model.Instances.New<IfcPropertySingleValue>(prop =>
                {
                    prop.Name = "YoungModulus";
                    prop.NominalValue = new IfcModulusOfElasticityMeasure(24.855578);
                    prop.Unit = modulusOfElasticityUnit;
                }));

            });
            #endregion

            #region Assigning The material to the Column
            var relAssociatesMaterial = model.Instances.New<IfcRelAssociatesMaterial>();

            relAssociatesMaterial.RelatingMaterial = materialProfileSetUsage.ForProfileSet.MaterialProfiles.First.Material;
            relAssociatesMaterial.RelatedObjects.Add(columnTypeRelation.RelatingType);
            relAssociatesMaterial.RelatedObjects.Add(column);
            #endregion


            //  Add to the storey
            //var storey = model.Instances.OfType<IfcBuildingStorey>().FirstOrDefault();
            //if (storey != null)
            storey.AddElement(column);

            //txn.Commit();
            return column;
        }

        //---------------------------------  Create IFC File  ----------------------------------//
        private static void CreateIFCFile(string fileName, string ifcfullPath, string xmlfullPath, List<Frame> framesList)
        {
            var model = CreateModelandAddProject(fileName + "Project");

            var site = CreateSite(model, fileName + "Site");

            var building = CreateBuilding(model, fileName + "Building");

            var storey1 = CreateStorey(model, fileName + "storey 1");

            //var column = CreateIfcColumn(model, 400, 300, 5000, 1000, 1000, 00);
            using (var txn = model.BeginTransaction("Create Columns"))
            {
                foreach (Frame frame in framesList)
                {
                    var column = CreateIfcColumn(model, storey1, frame.Width, frame.Height, frame.Xstart, frame.Ystart, frame.Zstart, frame.Xend, frame.Yend, frame.Zend);
                }
                txn.Commit();
            }

            /*Write the Ifc File with ifc extension*/
            model.SaveAs(ifcfullPath, IfcStorageType.Ifc);

            /*Write the Ifc File with ifcxml extension*/
            model.SaveAs(xmlfullPath, IfcStorageType.IfcXml);
        }

        private static void CreateIFCFile(string fileName, string ifcfullPath, string xmlfullPath, double length, double width, double height, List<Point> pointsList)
        {
            var model = CreateModelandAddProject(fileName + "Project");

            var site = CreateSite(model, fileName + "Site");

            var building = CreateBuilding(model, fileName + "Building");

            var storey1 = CreateStorey(model, fileName + "storey 1");

            //var column = CreateIfcColumn(model, 400, 300, 5000, 1000, 1000, 00);
            using (var txn = model.BeginTransaction("Create Columns"))
            {
                foreach (Point point in pointsList)
                {
                    var column = CreateIfcColumn(model, storey1, length, width, height, point.X, point.Y, point.Z);
                }
                txn.Commit();
            }

            /*Write the Ifc File with ifc extension*/
            model.SaveAs(ifcfullPath, IfcStorageType.Ifc);

            /*Write the Ifc File with ifcxml extension*/
            model.SaveAs(xmlfullPath, IfcStorageType.IfcXml);
        }

        //---------------------------------  Modify IFC File  ----------------------------------//
        public ActionResult Modify(int hiddenId, double newX, double newY)
        {
            var wexBimFileName = System.Configuration.ConfigurationManager.AppSettings["wexBIMFileName"];
            var stepBimFileName = Path.ChangeExtension(wexBimFileName, "ifc");
            string physicalPath = Path.Combine(Server.MapPath("~/data"), stepBimFileName);
            using (var model = IfcStore.Open(physicalPath))
            {
                var pickedColumn = model.Instances.Where(
                e => e.EntityLabel == hiddenId
                ).FirstOrDefault();

                if (pickedColumn is IIfcColumn)
                {
                    var castedColumn = pickedColumn as IfcColumn;

                    //var pickedColumn = model.Instances.Where(
                    //e => e.EntityLabel == hiddenId
                    //).FirstOrDefault() as IIfcColumn;

                    //var sweptArea = (castedColumn.Representation.Representations.First.Items.First as IIfcExtrudedAreaSolid).SweptArea as IIfcRectangleProfileDef;
                    using (var txn = model.BeginTransaction("Transaction Name"))
                    {

                        (castedColumn.Representation.Representations.First.Items.First as IIfcExtrudedAreaSolid).SweptArea = model.Instances.New<IfcRectangleProfileDef>(
                            rpd =>
                            {
                                rpd.ProfileName = $"{newX}x{newY}";
                                rpd.XDim = newX;
                                rpd.YDim = newY;
                                rpd.Position = model.Instances.New<IfcAxis2Placement2D>(a2p2d =>
                                {
                                    a2p2d.Location = model.Instances.New<IfcCartesianPoint>(cp =>
                                    {
                                        cp.SetXY(0, 0);
                                    });
                                    a2p2d.RefDirection = model.Instances.New<IfcDirection>(cp =>
                                    {
                                        cp.SetXY(1, 0);
                                    });
                                });
                            }
                            );

                        //sweptArea.ProfileName= $"{newX}x{newY}";
                        //sweptArea.XDim = newX;
                        //sweptArea.YDim = newY;
                        txn.Commit();
                        model.SaveAs(physicalPath);
                    }
                }

                ConvertToWexBIM(physicalPath);

            }

            return View("Viewer");
        }



        //---------------------------------  Read Excel File  ----------------------------------//
        private static List<Frame> ReadFromThreeeDimExcelFile(string filePath)
        {

            //List<List<double>> pointsList = new List<List<double>>();
            List<Frame> framesList = new List<Frame>();

            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.FirstOrDefault();
                object thisValue;
                Frame frame;
                for (int row = 2; row < 9999; row++)
                {
                    thisValue = sheet.Cells[row, 1].Value;
                    if (thisValue == null || thisValue.ToString() == "")
                        break;
                    frame = new Frame();
                    for (int col = 1; col < 9; col++)
                    {
                        thisValue = sheet.Cells[row, col].Value;
                        if (thisValue == null || thisValue.ToString() == "")
                            break;

                        switch (col)
                        {
                            case 1:
                                frame.Width = (double)thisValue;
                                break;
                            case 2:
                                frame.Height = (double)thisValue;
                                break;
                            case 3:
                                frame.Xstart = (double)thisValue;
                                break;
                            case 4:
                                frame.Ystart = (double)thisValue;
                                break;
                            case 5:
                                frame.Zstart = (double)thisValue;
                                break;
                            case 6:
                                frame.Xend = (double)thisValue;
                                break;
                            case 7:
                                frame.Yend = (double)thisValue;
                                break;
                            case 8:
                                frame.Zend = (double)thisValue;
                                break;
                        }
                    }
                    bool contains = false;
                    if (framesList.Count == 0)
                        framesList.Add(frame);
                    else
                    {
                        for (int i = 0; i < framesList.Count; i++)
                        {
                            if (frame.Xstart == framesList[i].Xstart && frame.Ystart == framesList[i].Ystart && frame.Zstart == framesList[i].Zstart && frame.Xend == framesList[i].Xend && frame.Yend == framesList[i].Yend && frame.Zend == framesList[i].Zend)
                                contains = true;
                        }
                        if (contains == false)
                            framesList.Add(frame);
                    }
                }
            }
            return framesList;
        }

        private static List<Point> ReadFromTwoDimExcelFile(string filePath)
        {
            //List<List<double>> pointsList = new List<List<double>>();
            List<Point> pointsList = new List<Point>();

            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.FirstOrDefault();
                object thisValue;
                Point point;
                for (int row = 2; row < 9999; row++)
                {
                    thisValue = sheet.Cells[row, 1].Value;
                    if (thisValue == null || thisValue.ToString() == "")
                        break;
                    point = new Point();
                    for (int col = 1; col < 4; col++)
                    {
                        thisValue = sheet.Cells[row, col].Value;
                        if (thisValue == null || thisValue.ToString() == "")
                            break;

                        switch (col)
                        {
                            case 1:
                                point.X = (double)thisValue;
                                break;
                            case 2:
                                point.Y = (double)thisValue;
                                break;
                            case 3:
                                point.Z = (double)thisValue;
                                break;
                        }
                    }

                    bool contains = false;
                    if (!pointsList.Contains(point))

                        if (pointsList.Count == 0)
                            pointsList.Add(point);
                        else
                        {
                            for (int i = 0; i < pointsList.Count; i++)
                            {
                                if (point.X == pointsList[i].X && point.Y == pointsList[i].Y && point.Z == pointsList[i].Z)
                                    contains = true;
                            }
                            if (contains == false)
                                pointsList.Add(point);
                        }

                }

            }

            return pointsList;
        }
        #endregion

    }

    public class Frame
    {
        public double Width { get; set; }
        public double Height { get; set; }
        public double Xstart { get; set; }
        public double Ystart { get; set; }
        public double Zstart { get; set; }
        public double Xend { get; set; }
        public double Yend { get; set; }
        public double Zend { get; set; }
    }
    public class Point
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
    }

}