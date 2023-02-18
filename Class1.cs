using System;
using System.Collections.Generic;
using NXOpen;
using NXOpen.Assemblies;
using NXOpen.UF;
using System.IO;
using NXOpen.Features;
using NXOpen.Tooling;
using NXOpen.CAE;

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

        public static Face SelectAnyFace()
        {
            Selection.SelectionType[] Types = { Selection.SelectionType.Faces };
            Selection.Response resp = theUI.SelectionManager.SelectObject("Select any Face", "Selection", Selection.SelectionScope.AnyInAssembly, false, Types, out NXObject nxObj, out Point3d cursor);
            if (resp == Selection.Response.ObjectSelected || resp == Selection.Response.ObjectSelectedByName)
            {
                return (Face)nxObj;
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
        public static void getHoles2(Part part,ref List<double> diaVales, ref List<double> depthValues, ref List<double> dirX, ref List<double> dirY, ref List<double> dirZ, ref List<Face> faces)
        {
            int edit = 0;
            string diameter = string.Empty;
            string depth = string.Empty;
            string tipAngle = string.Empty;
            int throughFlag = 0;

            try
            {

                foreach (Feature feature in part.Features)
                {
                    if (feature.FeatureType == "SIMPLE HOLE")
                    {
                        UF.Modl.AskSimpleHoleParms(feature.Tag, edit, out diameter, out depth, out tipAngle, out throughFlag);
                        string[] separator = { " ", "=" };
                        int count = 2;
                        string[] DiaValue = diameter.Split(separator, count, StringSplitOptions.RemoveEmptyEntries);
                        string[] DepValue = depth.Split(separator, count, StringSplitOptions.RemoveEmptyEntries);
                        for (int i = 0; i < DiaValue.Length; i++)
                        {
                            if (i % 2 != 0)
                            {
                                //Console.WriteLine(val1[i]);
                                diaVales.Add(Convert.ToDouble(DiaValue[i]));

                            }
                        }
                        for (int i = 0; i < DepValue.Length; i++)
                        {
                            if (i % 2 != 0)
                            {
                                //Console.WriteLine(val1[i]);
                                depthValues.Add(Convert.ToDouble(DepValue[i]));
                            }
                        }
                        BodyFeature hole = (BodyFeature)feature;
                        dirX.Add(hole.Location.X);
                        dirY.Add(hole.Location.Y);
                        dirZ.Add(hole.Location.Z);
                        Face[] f = hole.GetFaces();
                        foreach (Face fa in f)
                        {
                            faces.Add(fa);
                        }
                    }
                }

                //Echo($"{depthValues.Count}");
                //Echo($"{diameterValues.Count}");

                //Echo($"{faces.Count}");

                //int countofholes = depthValues.Count;


                //for (int i = 0; i < faces.Count; i++)
                //{
                //    Echo($"{faces[i].JournalIdentifier.ToString()}");
                //}

                //for (int b = 0; b < diameterValues.Count; b++)
                //{
                //    Echo($"{diameterValues[b]}");
                //}

                //int val = 0;
                //foreach (double b in diameterValues)
                //{
                //    DataForHole.Add(b, new DataForFastners(depthValues[val], directionX[val], directionY[val], directionZ[val], faces[val]));

                //}

                //for (int j = 0; j < diameterValues.Count; j++)
                //{
                //    //Echo($"{j}");
                //    DataForHole.Add(diameterValues[j], new DataForFastners(depthValues[j], directionX[j], directionY[j], directionZ[j], faces[j]));
                //}
            }
            catch (NXException ex)
            {
                theUI.NXMessageBox.Show("Get Holes", NXMessageBox.DialogType.Error, $"{ex.Message} {ex.StackTrace}");
            }
        }

        public static void FastnersAssembly(int a,double diameter, double depth, double X, double Y, double Z, Face f, Face[] RefrenceFces, Face BoltInserFace)
        {

            //Echo($"--------------");
            //Echo($"{diameter}");
            //Echo($"{depth}");
            //Echo($"--------------");
            int type1;
            double[] point1 = new double[3];
            double[] dir1 = new double[3];
            double[] box1 = new double[6];
            double radius1;
            double rad1;
            int norm_dir1;
            Part workPart = session.Parts.Work;
            try
            {
                
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
                    UF.Modl.AskFaceData(BoltInserFace.Tag, out type1, point1, dir1, box1, out radius1, out rad1, out norm_dir1);
                }
                else
                {
                    UF.Modl.AskFaceData(BoltInserFace.Tag, out type1, point1, dir1, box1, out radius1, out rad1, out norm_dir1);

                }
                //direction or point
                fastenerAssy1.SetHoleDirection(new Point3d(dir1[0], dir1[1], dir1[2]), 0);
                fastenerAssy1.SetHolePosition(new Point3d(X, Y, Z), 0);
                fastenerAssy1.SetHoleOriginPosition(new Point3d(X, Y, Z), 0);
                fastenerAssy1.SetHoleHeight(depth, 0);
                fastenerAssy1.SetHoleOriginHeight(depth, 0);
                fastenerAssy1.SetHoleOriginDiameter(diameter, 0);
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
                //fastenerAssy1.EraseFastenerAssemblyData(0);
                FastenerAssemConfigBuilder fastenerAssemConfigBuilder2 = workPart.ToolingManager.FastenerAssemConfig.CreateBuilder();
                //NXOpen.Gateway.ImageCaptureBuilder imageCaptureBuilder1 = workPart.ImageCaptureManager.CreateImageCaptureBuilder();
                //imageCaptureBuilder1.Size = NXOpen.Gateway.ImageCaptureBuilder.ImageSize.Pixels64;
                //imageCaptureBuilder1.Size = NXOpen.Gateway.ImageCaptureBuilder.ImageSize.Pixels128;
                //imageCaptureBuilder1.Format = NXOpen.Gateway.ImageCaptureBuilder.ImageFormat.Bmp;
                //imageCaptureBuilder1.ImageFile = "C:\\Users\\durga\\AppData\\Local\\Temp\\durg2380FF505kpv.bmp";
                fastenerAssemConfigBuilder2.Destroy();
                //imageCaptureBuilder1.Destroy();
                fastenerAssy1.AddParentNewPart("C:\\NX\\MajorProjects\\Parts\\GB-Hex Bolt Stacks"+$"{a}"+".prt", 0, true);
                NXObject nXObject1 = fastenerAssy1.AddTopNode(new Point3d(X, Y, Z), new Point3d(dir1[0], dir1[1], dir1[2]), f, 0);
                fastenerAssy1.AddScrewArray("ANSI Metric\\Bolt\\Hex Head\\Hex Bolt, AM.krx", "LENGTH", "C:\\Program Files\\Siemens\\NX2007\\nxparts\\Reuse Library\\Reuse Examples\\Standard Parts", "Fastener Assembly Configuration Library", "ANSI Metric\\Bolt\\Hex Head\\Hex Bolt, AM.prt", 0, NXOpen.Tooling.FastenerAssy.StackTypeMethod.Screw);

                fastenerAssy1.AddScrewArray("ANSI Metric\\Washer\\Plain\\Plain Washer, Regular, AM.krx", "THICKNESS", "C:\\Program Files\\Siemens\\NX2007\\nxparts\\Reuse Library\\Reuse Examples\\Standard Parts", "Fastener Assembly Configuration Library", "ANSI Metric\\Washer\\Plain\\Plain Washer, Regular, AM.prt", 0, NXOpen.Tooling.FastenerAssy.StackTypeMethod.TopStack);

                fastenerAssy1.AddScrewArray("ANSI Metric\\Washer\\Plain\\Plain Washer, Regular, AM.krx", "THICKNESS", "C:\\Program Files\\Siemens\\NX2007\\nxparts\\Reuse Library\\Reuse Examples\\Standard Parts", "Fastener Assembly Configuration Library", "ANSI Metric\\Washer\\Plain\\Plain Washer, Regular, AM.prt", 0, NXOpen.Tooling.FastenerAssy.StackTypeMethod.BottomStack);

                if (diameter <= 15)
                {
                    fastenerAssy1.AddScrewArray("ANSI Metric\\Nut\\Hex\\Hex Nut, Small, 1, AM.krx", "THICKNESS", "C:\\Program Files\\Siemens\\NX2007\\nxparts\\Reuse Library\\Reuse Examples\\Standard Parts", "Fastener Assembly Configuration Library", "ANSI Metric\\Nut\\Hex\\Hex Nut, Small, 1, AM.prt", 0, NXOpen.Tooling.FastenerAssy.StackTypeMethod.BottomStack);

                }
                else
                {
                    fastenerAssy1.AddScrewArray("ANSI Metric\\Nut\\Hex\\Hex Nut, Big, 1, AM.krx", "THICKNESS", "C:\\Program Files\\Siemens\\NX2007\\nxparts\\Reuse Library\\Reuse Examples\\Standard Parts", "Fastener Assembly Configuration Library", "ANSI Metric\\Nut\\Hex\\Hex Nut, Small, 1, AM.prt", 0, NXOpen.Tooling.FastenerAssy.StackTypeMethod.BottomStack);

                }

                fastenerAssy1.UpdateFastenerStacks(0, true, false);
                workPart = session.Parts.Work;
                fastenerAssy1.UpdateDefaultStandard(0, "Common", "Simple", "");
                fastenerAssy1.SetAssemblyExtentLength(0, 1.5);
                //fastenerAssy1.EraseFastenerAssemblyData(0);
                fastenerAssy1.DeleteArrayHole(0);
                fastenerAssy1.RemoveFastenerConstraints(0);
                fastenerAssy1.DeleteArrayHole(0);
                fastenerAssy1.SaveUdoData();
                fastenerAssy1.EraseFastenerSetupData();
                fastenerAssemConfigBuilder1.Destroy();
                session.CleanUpFacetedFacesAndEdges();
                
                
            }
            catch (NXException ex)
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
            try
            {
                program = new FlangeAutomateAssembly();
                CreateTemplateFile();
                InsertComponent(@"C:\NX\MajorProjects\Parts\Flange with simple holes.prt", @"Base Flange", 0, 0, 0);
                Component cmp1 = selectComponent();
                PartLoadStatus pls;
                session.Parts.SetWorkComponent(cmp1, PartCollection.RefsetOption.Entire, PartCollection.WorkComponentOption.Visible, out pls);
                Part workPart = session.Parts.Work;
                pls.Dispose();
                List<double> diaVales1 = new List<double>();
                List<double> depthValues1 = new List<double>();
                List<double> dirX1 = new List<double>();
                List<double> dirY1 = new List<double>();
                List<double> dirZ1 = new List<double>();
                List<Face> faceOfHole = new List<Face>();
                getHoles2(workPart, ref diaVales1, ref depthValues1, ref dirX1, ref dirY1, ref dirZ1, ref faceOfHole);
                Component nullNXOpen_Assemblies_Component = null;
                PartLoadStatus partLoadStatus2;
                session.Parts.SetWorkComponent(nullNXOpen_Assemblies_Component, PartCollection.RefsetOption.Entire, PartCollection.WorkComponentOption.Visible, out partLoadStatus2);
                workPart = session.Parts.Work; // FlangeAssembly1
                partLoadStatus2.Dispose();
                Part work = session.Parts.Work;
                Part display = session.Parts.Display;
                Face[] faceForMate = new Face[2];
                faceForMate[0] = SelectAnyFace();
                faceForMate[1] = SelectAnyFace();
                //Echo($"{faceForMate[1].JournalIdentifier}");
                int countOfAllData = diaVales1.Count;
                //foreach(double d in diaVales1)
                //{
                //    Echo($"Diameter {d}");
                //}
                //Echo($"{diaVales1.Count}");
                //Echo($"{depthValues1.Count}");
                //Echo($"{dirX1.Count}");
                //Echo($"{dirY1.Count}");
                //Echo($"{dirZ1.Count}");
                //Echo($"{faceOfHole.Count}");
                if (countOfAllData == depthValues1.Count && countOfAllData == dirX1.Count && countOfAllData == dirY1.Count && countOfAllData == dirZ1.Count && countOfAllData == faceOfHole.Count)
                {
                    for (int i = 0; i < countOfAllData; i++)
                    {
                        FastnersAssembly(i, diaVales1[i], depthValues1[i], dirX1[i], dirY1[i], dirZ1[i], faceOfHole[i], faceForMate, faceForMate[0]);
                    }
                }
                else
                {
                    Echo("not same");
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
            //huuuu cute fellow
        }
    }
}
