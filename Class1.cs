using System;
using System.Collections.Generic;
using NXOpen;
using NXOpen.Assemblies;
using NXOpen.UF;
using System.IO;
using NXOpen.Features;
using NXOpen.Tooling;

namespace FlangeAutomation
{
    public class FlangeAutomateAssembly
    {
        static Session session;
        static UFSession UF;
        static UI theUI;
        static ListingWindow lw;
        public static FlangeAutomateAssembly program;
        public static bool isDisposeCalled;

        public FlangeAutomateAssembly()
        {
            try
            {
                session = Session.GetSession();
                UF = UFSession.GetUFSession();
                theUI = UI.GetUI();
                lw = session.ListingWindow;
                isDisposeCalled = false;
            }
            catch(NXException ex)
            {
                theUI.NXMessageBox.Show("Constructor", NXMessageBox.DialogType.Error, $"{ex.Message} {ex.StackTrace}");
            }
        }

        public static Face[] SelectFaces()
        {
            Selection.SelectionType[] types = { Selection.SelectionType.Faces };
            Selection.Response resp = theUI.SelectionManager.SelectObjects("Select any two Face", "Select any two faces", Selection.SelectionScope.AnyInAssembly, false, types, out NXObject[] nxObj);
            if (resp == Selection.Response.ObjectSelected || resp == Selection.Response.ObjectSelectedByName)
            {
                return (Face[])nxObj;
            }
            else
            {
                return null;
            }
        }

        public static Component selectComponent()
        {

            Selection.MaskTriple[] mask = new Selection.MaskTriple[1];
            mask[0].Type = UFConstants.UF_component_type;
            mask[0].Subtype = 0;
            mask[0].SolidBodySubtype = 0;

            Selection.Response resp = theUI.SelectionManager.SelectObject("Seelect component", "Select component", Selection.SelectionScope.AnyInAssembly, Selection.SelectionAction.ClearAndEnableSpecific, false, false, mask, out NXObject nxObj, out Point3d cursor);
            if(resp == Selection.Response.ObjectSelected||resp == Selection.Response.ObjectSelectedByName)
            {
                return (Component)nxObj;
            }
            else
            {
                return null;
            }
        }

        public static void CreateTemplateFile()
        {
            try
            {
                FileNew createfile = session.Parts.FileNew();
                createfile.Units = Part.Units.Millimeters;
                createfile.TemplateFileName = "assembly-mm-template.prt";
                string filename = createfile.NewFileName = @"C:\NX\MajorProjects\Parts\FlangeAssembly1.prt";
                if (File.Exists(filename))
                {
                    int num = theUI.NXMessageBox.Show("Create Template", NXMessageBox.DialogType.Question, "File Name Already Exist...! Do u want to delete existing File and Create New file");
                    if (num == 1)
                    {
                        File.Delete(filename);
                    }
                }
                createfile.Commit();

            }
            catch(NXException ex)
            {
                theUI.NXMessageBox.Show("Create Template", NXMessageBox.DialogType.Error, $"{ex.Message} {ex.StackTrace}");

            }

        }

        public static void InsertComponent(string partToLoad, string compName,double X,double Y,double Z)
        {
            Part workPart = session.Parts.Work;
            try
            {
                Point3d origin = new Point3d(X, Y, Z);
                workPart.WCS.Rotate(WCS.Axis.XAxis, 0);
                string refSet = "MODEL";
                Matrix3x3 mat1 = workPart.WCS.CoordinateSystem.Orientation.Element;
                PartLoadStatus pls = null;

                workPart.ComponentAssembly.AddComponent(partToLoad,refSet,compName,origin,mat1,1,out pls);

            }catch(NXException ex)
            {
                theUI.NXMessageBox.Show("Insert Component", NXMessageBox.DialogType.Error, $"{ex.Message} {ex.StackTrace}");
                
            }
        }

        public static void getHoles(Part part, ref Dictionary<double, DataForFastners> DataForHole)
        {
            int edit = 0;
            string diameter = string.Empty;
            string depth = string.Empty;
            string tipAngle = string.Empty;
            int throughFlag = 0;
            List<double> diameterValues = new List<double>();
            List<double> depthValues = new List<double>();
            string[] DiaValue = null;
            string[] DepValue = null;
            Face[] faces = null;
            List<double> directionX = new List<double>();
            List<double> directionY = new List<double>();
            List<double> directionZ = new List<double>();

            try
            {
                foreach (Feature feature in part.Features)
                {
                    if (feature.FeatureType == "SIMPLE HOLE")
                    {
                        UF.Modl.AskSimpleHoleParms(feature.Tag, edit, out diameter, out depth, out tipAngle, out throughFlag);
                        string[] separator = { " ", "=" };
                        int count = 2;
                        DiaValue = diameter.Split(separator, count, StringSplitOptions.RemoveEmptyEntries);
                        DepValue = depth.Split(separator, count, StringSplitOptions.RemoveEmptyEntries);
                        BodyFeature hole = (BodyFeature)feature;
                        directionX.Add(hole.Location.X);
                        directionY.Add(hole.Location.Y);
                        directionZ.Add(hole.Location.Z);
                        faces = hole.GetFaces();
                    }
                }
                for (int a = 0; a < DiaValue.Length; a++)
                {
                    if (a % 2 != 0)
                    {
                        diameterValues.Add(Convert.ToDouble(DiaValue[a]));
                        depthValues.Add(Convert.ToDouble(DepValue[a]));
                    }
                }
                for(int b =0; b< diameterValues.Count; b++)
                {
                    DataForHole.Add(diameterValues[b], new DataForFastners(depthValues[b], directionX[b], directionY[b], directionZ[b], faces[b]));
                }
            }
            catch (NXException ex)
            {
                theUI.NXMessageBox.Show("Get Holes", NXMessageBox.DialogType.Error, $"{ex.Message} {ex.StackTrace}");
            }
        }

        public static void FastnersAssembly(Part workPart, double diameter, double depth, double X, double Y, double Z, Face f, Face[] RefrenceFces, Face BoltInserFace)
        {
            int type;
            double[] point = new double[3];
            double[] dir = new double[3];
            double[] box = new double[6];
            double radius;
            double rad;
            int norm_dir;

            try
            {
                workPart = session.Parts.Work;
                FastenerAssy fastenerAssy1 = workPart.ToolingManager.FastenerAssembly.CreateBuilder();
                FastenerAssemConfigBuilder fastenerAssemConfigBuilder1 = workPart.ToolingManager.FastenerAssemConfig.CreateBuilder();
                Sketch sketch1 = fastenerAssy1.PositioningFeature;
                fastenerAssy1.UpdateFastenerLength(true);
                fastenerAssy1.SetFastenerExtentLength(1.5);
                fastenerAssy1.SetFastenerSelectionType(FastenerAssy.SelectionTypeMethod.Hole);
                fastenerAssy1.SetFastenerExtentLength(1.5);
                fastenerAssy1.SetHoleDiameter(diameter, 0);
                if (BoltInserFace == RefrenceFces[0])
                {
                    UF.Modl.AskFaceData(BoltInserFace.Tag,out type,point,dir,box,out radius, out rad,out norm_dir);
                }
                else
                {
                    UF.Modl.AskFaceData(BoltInserFace.Tag, out type, point, dir, box, out radius, out rad, out norm_dir);
                }
                //direction or point
                fastenerAssy1.SetHoleDirection(new Point3d(dir[0], dir[1],dir[2]), 0);
                fastenerAssy1.SetHolePosition(new Point3d(X, Y, Z), 0);
                fastenerAssy1.SetHoleOriginPosition(new Point3d(X, Y, Z), 0);
                fastenerAssy1.SetHoleHeight(20.0, 0);
                fastenerAssy1.SetHoleOriginHeight(20.0, 0);
                fastenerAssy1.SetHoleOriginDiameter(50.0, 0);
                fastenerAssy1.SetHoleDefaultCylindricalFace(f, 0);
                fastenerAssy1.SetHoleSideCylindricalFaces(f, 0);
                fastenerAssy1.SetHoleSideCylindricalFaces(f, 0);
                fastenerAssy1.SetHoleFaces(f, 0);
                fastenerAssy1.SetSidePlanarFaces(RefrenceFces[0], 0);
                fastenerAssy1.SetSidePlanarFaces(RefrenceFces[1], 0);
                fastenerAssy1.SetDefaultPlanarFaces(RefrenceFces[0], 0);
                NXObject nullNXOpen_NXObject = null;
                fastenerAssy1.SetInstanceFeatureFaces(nullNXOpen_NXObject, 0);
                fastenerAssy1.SetFastenerMode(FastenerAssy.ModeMethod.Add);
                Component nullNXOpen_Assemblies_Component = null;
                fastenerAssy1.ReadAssemblyConfigure(0, nullNXOpen_Assemblies_Component);
                fastenerAssy1.DeleteArrayHole(0);
                fastenerAssy1.EraseFastenerAssembly(0, true, false, false, false, false, true, true, true);
                FastenerAssemConfigBuilder fastenerAssemConfigBuilder2 = workPart.ToolingManager.FastenerAssemConfig.CreateBuilder();
                NXOpen.Gateway.ImageCaptureBuilder imageCaptureBuilder1 = workPart.ImageCaptureManager.CreateImageCaptureBuilder();
                imageCaptureBuilder1.Size = NXOpen.Gateway.ImageCaptureBuilder.ImageSize.Pixels64;
                imageCaptureBuilder1.Size = NXOpen.Gateway.ImageCaptureBuilder.ImageSize.Pixels128;
                imageCaptureBuilder1.Format = NXOpen.Gateway.ImageCaptureBuilder.ImageFormat.Bmp;
                imageCaptureBuilder1.ImageFile = "C:\\Users\\durga\\AppData\\Local\\Temp\\durg2380FF505kpv.bmp";
                fastenerAssemConfigBuilder2.Destroy();
                imageCaptureBuilder1.Destroy();
                fastenerAssy1.AddParentNewPart("C:\\NX\\MajorProjects\\Parts\\GB-Hex Bolt Stacks 1.prt", 0, true);
                NXObject nXObject1 = fastenerAssy1.AddTopNode(new Point3d(X, Y, Z), new Point3d(dir[0], dir[1], dir[2]), f, 0);
                fastenerAssy1.AddScrewArray("ANSI Metric\\Bolt\\Hex Head\\Hex Bolt, AM.krx", "LENGTH", "C:\\Program Files\\Siemens\\NX2007\\nxparts\\Reuse Library\\Reuse Examples\\Standard Parts", "Fastener Assembly Configuration Library", "ANSI Metric\\Bolt\\Hex Head\\Hex Bolt, AM.prt", 0, NXOpen.Tooling.FastenerAssy.StackTypeMethod.Screw);
                fastenerAssy1.AddScrewArray("ANSI Metric\\Washer\\Plain\\Plain Washer, Regular, AM.krx", "THICKNESS", "C:\\Program Files\\Siemens\\NX2007\\nxparts\\Reuse Library\\Reuse Examples\\Standard Parts", "Fastener Assembly Configuration Library", "ANSI Metric\\Washer\\Plain\\Plain Washer, Regular, AM.prt", 0, NXOpen.Tooling.FastenerAssy.StackTypeMethod.TopStack);
                fastenerAssy1.AddScrewArray("ANSI Metric\\Nut\\Hex\\Hex Nut, Big, 1, AM.krx", "THICKNESS", "C:\\Program Files\\Siemens\\NX2007\\nxparts\\Reuse Library\\Reuse Examples\\Standard Parts", "Fastener Assembly Configuration Library", "ANSI Metric\\Nut\\Hex\\Hex Nut, Big, 1, AM.prt", 0, NXOpen.Tooling.FastenerAssy.StackTypeMethod.BottomStack);
                fastenerAssy1.UpdateFastenerStacks(0, true, false);
                fastenerAssy1.UpdateDefaultStandard(0, "Common", "Simple", "");
                fastenerAssy1.SetAssemblyExtentLength(0, 1.5);
                fastenerAssy1.DeleteArrayHole(0);
                fastenerAssy1.RemoveFastenerConstraints(0);
                fastenerAssy1.DeleteArrayHole(0);
                fastenerAssy1.SaveUdoData();
                fastenerAssy1.EraseFastenerSetupData();
                fastenerAssemConfigBuilder1.Destroy();
                session.CleanUpFacetedFacesAndEdges();
            }
            catch(NXException ex)
            {
                theUI.NXMessageBox.Show("Fastener Assembly", NXMessageBox.DialogType.Error, ex.Message + " " + ex.StackTrace);
            }
        }

        public static void Echo(string Output)
        {
            lw.Open();
            lw.WriteFullline(Output);
        }

        public static void Main()
        {
            Dictionary<double, DataForFastners> MyData = new Dictionary<double, DataForFastners>();
            
            try
            {
                program = new FlangeAutomateAssembly();
                CreateTemplateFile();
                InsertComponent(@"C:\NX\MajorProjects\Parts\RecFlange.prt", @"Base Flange", 0, 0, 0);
                Component cmp1 = selectComponent();
                Face[] face = SelectFaces();
                PartLoadStatus pls;
                session.Parts.SetWorkComponent(cmp1, PartCollection.RefsetOption.Entire, PartCollection.WorkComponentOption.Visible, out pls);
                pls.Dispose();
                Part workPart = session.Parts.Work;
                getHoles(workPart, ref MyData);
                int i =0;
                foreach (var items in MyData)
                {
                    FastnersAssembly(workPart, items.Key, items.Value.Depth, items.Value.XX, items.Value.YY, items.Value.ZZ, items.Value.Faces, face, face[0]);
                    i++;
                    if (i == 1)
                    {
                        break;
                    }
                }
                program.Dispose();

            }
            catch(NXException ex)
            {
                theUI.NXMessageBox.Show("Main Function", NXMessageBox.DialogType.Error, $"{ex.Message} {ex.StackTrace}");

            }
        }

        public void Dispose()
        {
            if(isDisposeCalled == false)
            {
            }
            isDisposeCalled = true;
        }

        public static int GetUnloadOption()
        {
            return (int)1;
        }
    }

    /// <summary>
    /// Class for data
    /// </summary>
    public class DataForFastners
    {
        public double Depth { get; set; }
        public double XX { get; set; }
        public double YY { get; set; }
        public double ZZ { get; set; }
        public Face Faces { get; set; }

        /// <summary>
        /// To store data
        /// </summary>
        /// <param name="Depth"></param>
        /// <param name="XX"></param>
        /// <param name="YY"></param>
        /// <param name="ZZ"></param>
        /// <param name="Faces"></param>
        public DataForFastners(double Depth, double XX, double YY, double ZZ, Face Faces)
        {
            this.Depth = Depth;
            this.XX = XX;
            this.YY = YY;
            this.ZZ = ZZ;
            this.Faces = Faces;
        }
    }
}
