using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using Autodesk.Revit.ApplicationServices;
using Application = Autodesk.Revit.ApplicationServices.Application;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Reflection;
using Org.BouncyCastle.Math.EC;
using System.Net;
using NPOI.POIFS.FileSystem;
using Org.BouncyCastle.Utilities;
using System.Windows.Forms.VisualStyles;
using System.Collections.ObjectModel;
using Autodesk.Revit.UI.Selection;
using NPOI.SS.Formula.Functions;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;
using Floor = Autodesk.Revit.DB.Floor;
using System.Xml.Linq;
using Autodesk.Revit.DB.Architecture;
using NPOI.Util;
using System.Collections;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar;
using System.Windows.Forms;
using RibbonPanel = Autodesk.Revit.UI.RibbonPanel;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

using System.Windows.Media.Imaging;
using System.Windows.Documents;
using System.Windows.Media.Media3D;
using TextElement = Autodesk.Revit.DB.TextElement;
using NPOI.SS.Formula.Eval;
using Autodesk.Windows;
using NPOI.OpenXmlFormats.Dml.Chart;
using static System.Net.Mime.MediaTypeNames;
using System.Windows.Controls;
using Frame = Autodesk.Revit.DB.Frame;
using System.Text.RegularExpressions;
using System.Windows;
using MessageBox = System.Windows.Forms.MessageBox;
using Autodesk.Revit.DB;

namespace ICAP
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Startup : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
        public Result OnStartup(UIControlledApplication application)
        {
            RibbonTab Tab = null;
            Autodesk.Windows.RibbonControl ribbonControl = Autodesk.Windows.ComponentManager.Ribbon;

            String tapname = "Bim Engine".ToUpper();
            String ribbonpanel = "Import Cap";
            foreach (Autodesk.Windows.RibbonTab ribbonTab in ribbonControl.Tabs)
            {
                if (ribbonTab.Title == "Bim Engine".ToUpper())
                {
                    Tab = ribbonTab;
                }
            }
            if (Tab == null)
            {
                try
                {
                    application.CreateRibbonTab(tapname.ToUpper());
                }
                catch (Exception ex) { }
            }
            RibbonPanel panel = null;

            List<RibbonPanel> panals = application.GetRibbonPanels(tapname);

            foreach (RibbonPanel pnls in panals)
            {
                if (pnls.Name == ribbonpanel)
                {
                    panel = pnls;

                }
            }

            if (panel == null)
            {
                panel = application.CreateRibbonPanel(tapname, ribbonpanel);
            }

            importcap(panel, "Import Cap", "ICAP", "ICAP.bundle", "ICAP", "Import Cap");
            newimportcap(panel, "New Import Cap", "New ICAP", "ICAP.bundle", "ICAP", "Import Cap Version 2.0");
            return Result.Succeeded;
        }
        private void importcap(RibbonPanel panel, string name, string text, string bundlename, string dllname, string tooltip)
        {
            PushButton pushButton1 = panel.AddItem(new PushButtonData(name, text, @"C:\ProgramData\Autodesk\ApplicationPlugins\" + bundlename + @"\Contents\" + dllname + ".dll", "ICAP.Oldstartup")) as PushButton;
            pushButton1.ToolTip = tooltip;
            ContextualHelp contextHelp = new ContextualHelp(ContextualHelpType.Url, "https://www.capbim.com");
            pushButton1.SetContextualHelp(contextHelp);
            pushButton1.LargeImage = new BitmapImage(new Uri(@"C:\ProgramData\Autodesk\ApplicationPlugins\\ICAP.bundle\Contents\Icons\CAP.png"));
        }
        private void newimportcap(RibbonPanel panel, string name, string text, string bundlename, string dllname, string tooltip)
        {
            PushButton pushButton1 = panel.AddItem(new PushButtonData(name, text, @"C:\ProgramData\Autodesk\ApplicationPlugins\" + bundlename + @"\Contents\" + dllname + ".dll", "ICAP.Newstartup")) as PushButton;
            pushButton1.ToolTip = tooltip;
            ContextualHelp contextHelp = new ContextualHelp(ContextualHelpType.Url, "https://www.capbim.com");
            pushButton1.SetContextualHelp(contextHelp);
            pushButton1.LargeImage = new BitmapImage(new Uri(@"C:\ProgramData\Autodesk\ApplicationPlugins\\ICAP.bundle\Contents\Icons\CAP.png"));
        }
    }


    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Oldstartup : IExternalCommand
    {
        private string caplocation;
        private string schfamilytype;
        public string[] myStrings;
        public string Thickness = null;
        private string levl;
        private double _Wid;
        private string Fwidth;
        private string FOffset;
        private double w;
        private double _nativeWitdth;
        private double _floorTypeWid;
        public FilteredElementCollector levelfillRegions;
        public Level _l;
        private string lname;
        private string strt;
        private string[] myStr;
        private double radius;
        private XYZ center;
        private FamilyInstance instance;
        private XYZ doorxyz;
        private double _dw;
        private double _dh;
        private double _wlnth;

        ///public List<String> duplicatLevel = new List<String>();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            UIApplication uiapp = commandData.Application;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            var path = doc.PathName;
            var filepath = path.Replace(".rvt", ".cap");
            Autodesk.Revit.DB.View view = doc.ActiveView;
            #region Collect elements
            FilteredElementCollector collector = new FilteredElementCollector(doc);

            ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);

            IList<Element> Walls = collector.WherePasses(filter).WhereElementIsNotElementType().ToElements();

            ICollection<Element> doors = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).ToElements();

            ICollection<Element> Windows = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).ToElements();

            FilteredElementCollector collectorbeams = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType();

            FilteredElementCollector COLLSLAB = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType();

            FilteredElementCollector fillRegions = new FilteredElementCollector(doc).OfClass(typeof(FilledRegion));

            var colllevel = new FilteredElementCollector(doc).OfClass(typeof(Level)).OfCategory(BuiltInCategory.OST_Levels);

            IList<Element> Wallstype = collector.WherePasses(filter).ToElements();

            #endregion Collect elements
            System.Xml.XmlWriterSettings settings = new System.Xml.XmlWriterSettings();
            settings.Indent = true;
            settings.IndentChars = "    ";
            settings.OmitXmlDeclaration = false;
            settings.Encoding = System.Text.Encoding.UTF8;
            #region xml
            //using (System.Xml.XmlWriter writer = System.Xml.XmlWriter.Create(filepath, settings))
            //{
            //    writer.WriteStartDocument();
            //    writer.WriteStartElement("Export");
            //    writer.WriteStartElement("CAP");
            //    //Exportbeam(writer, collectorbeams, app, doc);
            //    // Exportfloor(writer, COLLSLAB, app, doc);
            //    // ExportFilledRegion(writer, fillRegions, view);
            //    //xportWalls(writer, Walls);
            //    //xportDoor(writer, doors);
            //    //XportWindow(writer, Windows);
            //    writer.Flush();
            //    writer.Close();
            //}
            #endregion

            #region open Dilog
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = "Browse Cap Files",
                CheckFileExists = true,
                CheckPathExists = true,
                DefaultExt = "cap",
                Filter = "Cap Files (*.cap)|*.cap",
                FilterIndex = 2,
                RestoreDirectory = true,
                ReadOnlyChecked = true,
                ShowReadOnly = true
            };
            #endregion open Dilog
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                caplocation = openFileDialog1.FileName;
                //  DecFile(caplocation);
                #region i cap
                string first = string.Empty; string second = string.Empty; string _Height = string.Empty; string _Volume = string.Empty; string _Width = string.Empty; string _Function = string.Empty;
                string Cfirst = string.Empty; string Csecond = string.Empty; string C_Height = string.Empty; string C_Volume = string.Empty; string C_Width = string.Empty; string CWLevel = string.Empty;

                string Wposx = string.Empty;
                string Wposy = string.Empty;
                string Wposz = string.Empty;
                string DoorName = string.Empty;
                string FStartPoint = string.Empty;
                string FEndPoint = string.Empty;
                string FLevel = string.Empty;
                string WLevel = string.Empty;
                string wfamil = string.Empty;
                string Fwidth = string.Empty;
                string FOffset = string.Empty;

                string FSStartPoint = string.Empty;
                string FSEndPoint = string.Empty;
                string SFLevel = string.Empty;
                string SFwidth = string.Empty;
                string SFOffset = string.Empty;

                string ColumnLevel = string.Empty;

                string FillregionPoint = string.Empty;
                string CircleFilledRegion = string.Empty;
                string Circlepoints = string.Empty;
                string Midpoint = string.Empty;
                string FillLevel = string.Empty;
                string BeamStartPoint = string.Empty;
                string BeamEndPoint = string.Empty;
                string BeamLevel = string.Empty;
                string BeamWidth = string.Empty;
                string BeamHight = string.Empty;

                string doorx = string.Empty;
                string doory = string.Empty;
                string doorz = string.Empty;
                string doorlevel = string.Empty;
                string doorwidth = string.Empty;
                string doorheight = string.Empty;
                string walllength = string.Empty;


                int n = 1;

                List<string> points = new List<string>();
                IList<String[]> Floorcoll = new List<String[]>();
                IList<String[]> SFloorcoll = new List<String[]>();

                IList<String[]> wallcoll = new List<String[]>();
                IList<String[]> Cwallcoll = new List<String[]>();
                IList<String[]> Fillregio = new List<String[]>();
                IList<String[]> circleFillregio = new List<String[]>();
                IList<String[]> Beamcoll = new List<String[]>();
                IList<String[]> doorcoll = new List<String[]>();

                #region c
                //                string caploction = Path.Combine(Path.GetFullPath(System.Reflection.Assembly.GetExecutingAssembly().Location)
                //.Remove(Path.GetFullPath(System.Reflection.Assembly.GetExecutingAssembly().Location).Count()
                //- Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location).Count()), @"ECAP\2D-3D.cap");

                //using (XmlReader reader = XmlReader.Create(caplocation))
                #endregion c
                using (XmlReader reader = XmlReader.Create(caplocation))
                {
                    while (reader.Read())
                    {
                        if (reader.IsStartElement())
                        {
                            //return only when you have START tag
                            switch (reader.Name.ToString())
                            {
                                //FilledRegion
                                case "Points":
                                    string Fillpoint = reader.ReadString();
                                    FillregionPoint = Fillpoint;
                                    break;

                                case "Circlepoints":
                                    string Cpoints = reader.ReadString();
                                    Circlepoints = Cpoints;
                                    break;

                                case "Radius":
                                    string Mpoint = reader.ReadString();
                                    Midpoint = Mpoint;
                                    break;
                                case "FillLevel":
                                    string Flleel = reader.ReadString();
                                    FillLevel = Flleel;
                                    break;


                                //Floor
                                case "FStartPoint":
                                    string Bsp = reader.ReadString();
                                    FStartPoint = Bsp;
                                    break;
                                case "FEndPoint":
                                    string Bep = reader.ReadString();
                                    FEndPoint = Bep;
                                    break;
                                case "Level":
                                    string lvl = reader.ReadString();
                                    FLevel = lvl;
                                    break;
                                case "FWidth":
                                    string Fwid = reader.ReadString();
                                    Fwidth = Fwid;
                                    break;
                                case "Offset":
                                    string ofset = reader.ReadString();
                                    FOffset = ofset;
                                    break;


                                case "FSStartPoint":
                                    string SBsp = reader.ReadString();
                                    FSStartPoint = SBsp;
                                    break;
                                case "FSEndPoint":
                                    string SBep = reader.ReadString();
                                    FSEndPoint = SBep;
                                    break;
                                case "SLevel":
                                    string Slvl = reader.ReadString();
                                    SFLevel = Slvl;
                                    break;
                                case "SFWidth":
                                    string SFwid = reader.ReadString();
                                    SFwidth = SFwid;
                                    break;
                                case "SOffset":
                                    string Sofset = reader.ReadString();
                                    SFOffset = Sofset;
                                    break;


                                //Wall

                                case "StartPoint":
                                    string startpoint = reader.ReadString();
                                    first = startpoint;
                                    break;
                                case "EndPoint":
                                    string endpoint = reader.ReadString();
                                    second = endpoint;
                                    break;
                                case "Height":
                                    string Height = reader.ReadString();
                                    _Height = Height;
                                    break;

                                case "Width":
                                    string Width = reader.ReadString();
                                    _Width = Width;
                                    break;
                                case "WLevel":
                                    string Wlvl = reader.ReadString();
                                    WLevel = Wlvl;
                                    break;
                                case "Family":
                                    string wallFamily = reader.ReadString();
                                    wfamil = wallFamily;
                                    break;

                                //C_Wall

                                case "CStartPoint":
                                    string Cstartpoint = reader.ReadString();
                                    Cfirst = Cstartpoint;
                                    break;
                                case "CEndPoint":
                                    string Cendpoint = reader.ReadString();
                                    Csecond = Cendpoint;
                                    break;
                                case "CHeight":
                                    string CHeight = reader.ReadString();
                                    C_Height = CHeight;
                                    break;

                                case "CWidth":
                                    string CWidth = reader.ReadString();
                                    C_Width = CWidth;
                                    break;
                                case "CWLevel":
                                    string CWlvl = reader.ReadString();
                                    CWLevel = CWlvl;
                                    break;



                                //Beam
                                case "BeamStartPoint":
                                    string Bemstr = reader.ReadString();
                                    BeamStartPoint = Bemstr;
                                    break;
                                case "BeamEndPoint":
                                    string Bemend = reader.ReadString();
                                    BeamEndPoint = Bemend;
                                    break;
                                case "BeamLevel":
                                    string Bemlevel = reader.ReadString();
                                    BeamLevel = Bemlevel;
                                    break;
                                case "BeamWidth":
                                    string BeamWid = reader.ReadString();
                                    BeamWidth = BeamWid;
                                    break;
                                case "BeamHight":
                                    string BHight = reader.ReadString();
                                    BeamHight = BHight;
                                    break;


                                //Door
                                case "X":
                                    string x = reader.ReadString();
                                    doorx = x;
                                    break;
                                case "Y":
                                    string Y = reader.ReadString();
                                    doory = Y;
                                    break;
                                case "Z":
                                    string z = reader.ReadString();
                                    doorz = z;
                                    break;
                                case "DoorLevel":
                                    string DLevel = reader.ReadString();
                                    doorlevel = DLevel;
                                    break;
                                case "Walllength":
                                    string wlen = reader.ReadString();
                                    walllength = wlen;
                                    break;


                                case "DoorWidth":
                                    string Dwid = reader.ReadString();
                                    doorwidth = Dwid;
                                    break;
                                case "DoorHeight":
                                    string Dhig = reader.ReadString();
                                    doorheight = Dhig;
                                    break;

                            }

                            if (FStartPoint.Length > 1 && FEndPoint.Length > 1 && FLevel.Length > 1 && Fwidth.Length > 0)
                            {
                                myStrings = new[] { FStartPoint, FEndPoint, FLevel, Fwidth, FOffset };
                                Floorcoll.Add(myStrings);
                                FStartPoint = string.Empty;
                                FEndPoint = string.Empty;
                                FLevel = string.Empty;
                                Fwidth = string.Empty;
                                FOffset = string.Empty;
                            }

                            if (FSStartPoint.Length > 1 && FSEndPoint.Length > 1 && SFLevel.Length > 1 && SFwidth.Length > 1)
                            {
                                myStrings = new[] { FSStartPoint, FSEndPoint, SFLevel, SFwidth, SFOffset };
                                SFloorcoll.Add(myStrings);
                                FSStartPoint = string.Empty;
                                FSEndPoint = string.Empty;
                                SFLevel = string.Empty;
                                SFwidth = string.Empty;
                                SFOffset = string.Empty;
                            }

                            if (first.Length > 1 && second.Length > 1 && _Height.Length > 1 && WLevel.Length > 1 && _Width.Length > 1)
                            {
                                myStrings = new[] { first, second, WLevel, _Height, _Width };
                                wallcoll.Add(myStrings);
                                first = string.Empty; second = string.Empty; _Height = string.Empty; _Volume = string.Empty; _Width = string.Empty; _Function = string.Empty; WLevel = string.Empty;
                            }

                            if (Cfirst.Length > 1 && Csecond.Length > 1 && C_Height.Length > 1 && CWLevel.Length > 1)
                            {
                                myStrings = new[] { Cfirst, Csecond, CWLevel, C_Height, C_Width };
                                Cwallcoll.Add(myStrings);
                                Cfirst = string.Empty; Csecond = string.Empty; C_Height = string.Empty; C_Volume = string.Empty; C_Width = string.Empty; CWLevel = string.Empty;
                            }

                            if (FillregionPoint.Length > 1)
                            {

                                myStrings = new[] { FillregionPoint };
                                Fillregio.Add(myStrings);

                                if (Fillregio.Count() == 5)
                                {
                                    createfillregion(doc, Fillregio);
                                    Fillregio.Clear();
                                    FillregionPoint = string.Empty;
                                    ColumnLevel = string.Empty;
                                }
                            }

                            if (Circlepoints.Length > 1 && Midpoint.Length > 1 && FillLevel.Length > 1)
                            {

                                myStrings = new[] { Circlepoints, Midpoint, FillLevel };
                                circleFillregio.Add(myStrings);

                                if (circleFillregio.Count() == 1)
                                {
                                    createCirclefillregion(doc, circleFillregio);
                                    circleFillregio.Clear();
                                    Midpoint = string.Empty;
                                    FillLevel = string.Empty;
                                }
                            }

                            if (BeamStartPoint.Length > 1 && BeamEndPoint.Length > 1 && BeamLevel.Length > 1 && BeamWidth.Length > 1 && BeamHight.Length > 1)
                            {
                                myStrings = new[] { BeamStartPoint, BeamEndPoint, BeamLevel, BeamWidth, BeamHight };
                                Beamcoll.Add(myStrings);
                                BeamStartPoint = string.Empty; BeamEndPoint = string.Empty; BeamLevel = string.Empty; BeamWidth = string.Empty; BeamHight = string.Empty;
                            }

                            if (doorx.Length > 1 && doory.Length > 1 && doorz.Length > 0 && doorlevel.Length > 1 && walllength.Length > 1 && doorwidth.Length > 1 && doorheight.Length > 1)
                            {
                                myStrings = new[] { doorx, doory, doorz, doorlevel, walllength, doorwidth, doorheight };
                                doorcoll.Add(myStrings);

                                doorx = string.Empty;
                                doory = string.Empty;
                                doorz = string.Empty;
                                doorlevel = string.Empty;
                                walllength = string.Empty;
                                doorwidth = string.Empty;
                                doorheight = string.Empty;
                            }

                        }
                    }
                }
                #endregion i cap

                #region Column

                foreach (var lvlname in colllevel)
                {
                    Level _l = (lvlname) as Level;

                    createColumn(doc, _l);
                }

                MessageBox.Show("Column Placed", "2D-3D");
                #endregion Column

                #region Floor
                foreach (var lvlname in colllevel)
                {
                    Level _l = (lvlname) as Level;

                    if (Floorcoll.Count > 1)
                    {
                        createfloor(doc, _l, Floorcoll, FLevel);
                        FEndPoint = string.Empty; FEndPoint = string.Empty; Fwidth = string.Empty; FOffset = string.Empty;
                    }
                    //if (SFloorcoll.Count > 1)
                    //{
                    //    createfloor(doc, _l, SFloorcoll, SFLevel);
                    //    FSEndPoint = string.Empty; FSEndPoint = string.Empty; SFwidth = string.Empty; SFOffset = string.Empty;

                    //}

                }
                MessageBox.Show("Floor Placed", "2D-3D");
                #endregion Floor

                #region Beam

                foreach (var lvlname in colllevel)
                {
                    Level _l = (lvlname) as Level;
                    createBeam(Beamcoll, _l, doc);
                    BeamStartPoint = string.Empty; BeamEndPoint = string.Empty; BeamLevel = string.Empty; BeamWidth = string.Empty; BeamHight = string.Empty;
                }
                MessageBox.Show("Beam Placed", "2D-3D");

                #endregion Beam

                #region Wall
                foreach (var lvlname in colllevel)
                {
                    Level _l = (lvlname) as Level;
                    if (wallcoll.Count > 1)
                    {
                        createwall(wallcoll, first, second, doc, _l, WLevel);
                        first = string.Empty; second = string.Empty; _Height = string.Empty; _Volume = string.Empty; _Width = string.Empty; _Function = string.Empty; WLevel = string.Empty;
                    }
                    //if (Cwallcoll.Count > 1)
                    //{
                    //    createCwall(Cwallcoll, Cfirst, Csecond, doc, _l, CWLevel);
                    //    Cfirst = string.Empty; Csecond = string.Empty; C_Height = string.Empty; C_Volume = string.Empty; C_Width = string.Empty; CWLevel = string.Empty;
                    //}

                }

                MessageBox.Show("Wall Placed", "2D-3D");
                #endregion Wall
                #region Roof
                foreach (var lvlname in colllevel)
                {
                    Level _l = (lvlname) as Level;
                    createroof(doc, _l);
                }


                MessageBox.Show("Roof Placed", "2D-3D");
                #endregion Roof
                #region Door
                //foreach (var lvlname in colllevel)
                //{
                //    Level _l = (lvlname) as Level;
                //    createDoor(doorcoll, _l, doc);
                //    doorx = string.Empty; doory = string.Empty; doorz = string.Empty; doorlevel = string.Empty; walllength = string.Empty; doorwidth = string.Empty;
                //    doorheight = string.Empty;

                //}
                //MessageBox.Show("Door Placed", "2D-3D");
                #endregion Door

                TaskDialog.Show("2D-3D", "ICAP Completed");
            }
            //EncFile(caplocation);
            return Result.Succeeded;
        }
        private void createroof(Document doc, Level _l)
        {
            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).Where<Element>(e => !string.IsNullOrEmpty(e.Name) && e.Name.Equals("Roof")).FirstOrDefault<Element>() as Level;
            RoofType roofType = new FilteredElementCollector(doc).OfClass(typeof(RoofType)).FirstOrDefault<Element>() as RoofType;

            FilteredElementCollector collectorr = new FilteredElementCollector(doc);
            collectorr = new FilteredElementCollector(doc);
            collectorr.OfClass(typeof(RoofType));
            RoofType roofTypee = collectorr.Last() as RoofType;

            Application application = doc.Application;

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Wall)); Wall wall = null;

            List<Wall> walls = new List<Wall>();
            CurveArray footprint = application.Create.NewCurveArray();

            XYZ ptt1 = new XYZ(74.059211967, 71.466247122, 113.549868766);
            XYZ ptt2 = new XYZ(91.773670348, 53.751788740, 113.549868766);
            XYZ ptt3 = new XYZ(45.020940943, 6.999059335, 113.549868766);
            XYZ ptt4 = new XYZ(28.232728744, 25.639763900, 113.549868766);

            CurveArray profile1 = null;
            profile1 = new CurveArray();

            profile1.Append(Line.CreateBound(ptt1, ptt2));
            profile1.Append(Line.CreateBound(ptt2, ptt3));
            profile1.Append(Line.CreateBound(ptt3, ptt4));
            profile1.Append(Line.CreateBound(ptt4, ptt1));

            XYZ pt1 = new XYZ(-74.588603848, 72.909804899, 113.549868766);
            XYZ pt2 = new XYZ(-48.210651092, 72.909804899, 113.549868766);
            XYZ pt3 = new XYZ(-48.210651092, -3.697806649, 113.549868766);
            XYZ pt4 = new XYZ(-74.588603848, -3.697806649, 113.549868766);


            CurveArray profile = null;
            profile = new CurveArray();

            profile.Append(Line.CreateBound(pt1, pt2));
            profile.Append(Line.CreateBound(pt2, pt3));
            profile.Append(Line.CreateBound(pt3, pt4));
            profile.Append(Line.CreateBound(pt4, pt1));



            XYZ p1 = new XYZ(-36.402994529, -72.810446098, 113.549868766);
            XYZ p2 = new XYZ(-59.655819028, -78.547297744, 113.549868766);
            XYZ p3 = new XYZ(-64.907050038, -57.262809210, 113.549868766);
            XYZ p4 = new XYZ(-41.668340058, -51.468748062, 113.549868766);


            CurveArray profile2 = null;
            profile2 = new CurveArray();

            profile2.Append(Line.CreateBound(p1, p2));
            profile2.Append(Line.CreateBound(p2, p3));
            profile2.Append(Line.CreateBound(p3, p4));
            profile2.Append(Line.CreateBound(p4, p1));


            XYZ pp1 = new XYZ(-41.665956026, -51.137473043, 113.549868766);
            XYZ pp2 = new XYZ(-64.918780524, -56.874324689, 113.549868766);
            XYZ pp3 = new XYZ(-70.170011534, -35.589836155, 113.549868766);
            XYZ pp4 = new XYZ(-46.931301555, -29.795775007, 113.549868766);


            CurveArray profile3 = null;
            profile3 = new CurveArray();

            profile3.Append(Line.CreateBound(pp1, pp2));
            profile3.Append(Line.CreateBound(pp2, pp3));
            profile3.Append(Line.CreateBound(pp3, pp4));
            profile3.Append(Line.CreateBound(pp4, pp1));




            //CurveArray profile4 = null;
            //profile4 = new CurveArray();

            //profile4.Append(Line.CreateBound(Rp1, Rp2));
            //profile4.Append(Line.CreateBound(Rp2, Rp3));
            //profile4.Append(Line.CreateBound(Rp3, Rp4));
            //profile4.Append(Line.CreateBound(Rp4, Rp1));

            //XYZ Rpp1 = new XYZ(-74.785454242, 57.619370757, 113.549868766);
            //XYZ Rpp2 = new XYZ(-74.785454242, 34.482967629, 113.549868766);
            //XYZ Rpp3 = new XYZ(-91.727704740, 34.482967629, 113.549868766);
            //XYZ Rpp4 = new XYZ(-91.727704740, 57.619370757, 113.549868766);

            //CurveArray profile5 = null;
            //profile5 = new CurveArray();

            //profile5.Append(Line.CreateBound(Rpp1, Rpp2));
            //profile5.Append(Line.CreateBound(Rpp2, Rpp3));
            //profile5.Append(Line.CreateBound(Rpp3, Rpp4));
            //profile5.Append(Line.CreateBound(Rpp4, Rpp1));

            //XYZ RRpp1 = new XYZ(-16.268387155, -19.983889225, 113.549868766);
            //XYZ RRpp2 = new XYZ(-36.609594504, -19.983889225, 113.549868766);
            //XYZ RRpp3 = new XYZ(-36.609594504, 0.354414930, 113.549868766);
            //XYZ RRpp4 = new XYZ(-16.268387155, 0.354414930, 113.549868766);

            //CurveArray profile6 = null;
            //profile6 = new CurveArray();

            //profile6.Append(Line.CreateBound(RRpp1, RRpp2));
            //profile6.Append(Line.CreateBound(RRpp2, RRpp3));
            //profile6.Append(Line.CreateBound(RRpp3, RRpp4));
            //profile6.Append(Line.CreateBound(RRpp4, RRpp1));

            //XYZ RRppp1 = new XYZ(-16.268387155, -19.983889225, 113.549868766);
            //XYZ RRppp2 = new XYZ(-36.609594504, -19.983889225, 113.549868766);
            //XYZ RRppp3 = new XYZ(-36.609594504, 0.354414930, 113.549868766);
            //XYZ RRppp4 = new XYZ(-16.268387155, 0.354414930, 113.549868766);

            //CurveArray profile7 = null;
            //profile7 = new CurveArray();

            //profile7.Append(Line.CreateBound(RRppp1, RRppp2));
            //profile7.Append(Line.CreateBound(RRppp2, RRppp3));
            //profile7.Append(Line.CreateBound(RRppp3, RRppp4));
            //profile7.Append(Line.CreateBound(RRppp4, RRppp1));

            UIDocument uidoc = new UIDocument(doc);
            var colllevel = new FilteredElementCollector(doc).OfClass(typeof(Level)).OfCategory(BuiltInCategory.OST_Levels);
            ModelCurveArray footPrintToModelCurveMapping = new ModelCurveArray();
            using (Transaction transaction = new Transaction(doc, "Place Roof"))
            {
                if (!transaction.HasStarted()) transaction.Start();

                try
                {
                    if (_l.Name == "Level 3")
                    {
                        FootPrintRoof footprintRoof = doc.Create.NewFootPrintRoof(profile1, _l, roofTypee, out footPrintToModelCurveMapping);

                        foreach (ModelCurve curve in footPrintToModelCurveMapping)
                        {
                            //if (curve.GeometryCurve.Length > 22)
                            //{
                            footprintRoof.set_DefinesSlope(curve, true);
                            footprintRoof.set_SlopeAngle(curve, 0.6);
                            //footprintRoof.set_Offset(curve, -1300);
                            Parameter BO = footprintRoof.LookupParameter("Base Offset From Level");
                            BO.SetValueString("-6");
                            //}
                        }
                    }
                    if (_l.Name == "Level 1")
                    {
                        FootPrintRoof footprintRoof = doc.Create.NewFootPrintRoof(profile, _l, roofTypee, out footPrintToModelCurveMapping);
                        foreach (ModelCurve curve in footPrintToModelCurveMapping)
                        {
                            //if (curve.GeometryCurve.Length > 35)
                            //{
                            footprintRoof.set_DefinesSlope(curve, true);
                            footprintRoof.set_SlopeAngle(curve, 0.6);
                            //footprintRoof.set_Offset(curve, -1300);
                            Parameter BO = footprintRoof.LookupParameter("Base Offset From Level");
                            BO.SetValueString("16.000");
                            //}
                        }

                    }

                    if (_l.Name == "Level 2")
                    {
                        FootPrintRoof footprintRoof = doc.Create.NewFootPrintRoof(profile2, _l, roofTypee, out footPrintToModelCurveMapping);
                        foreach (ModelCurve curve in footPrintToModelCurveMapping)
                        {
                            if (curve.GeometryCurve.Length > 22)
                            {
                                footprintRoof.set_DefinesSlope(curve, true);
                                footprintRoof.set_SlopeAngle(curve, 0.6);
                                //footprintRoof.set_Offset(curve, -1300);
                                Parameter BO = footprintRoof.LookupParameter("Base Offset From Level");
                                BO.SetValueString("6");
                            }
                        }
                    }
                    if (_l.Name == "Ground")
                    {
                        FootPrintRoof footprintRooff = doc.Create.NewFootPrintRoof(profile3, _l, roofTypee, out footPrintToModelCurveMapping);
                        foreach (ModelCurve curve in footPrintToModelCurveMapping)
                        {
                            if (curve.GeometryCurve.Length > 22)
                            {
                                footprintRooff.set_DefinesSlope(curve, true);
                                footprintRooff.set_SlopeAngle(curve, 0.6);
                                //footprintRoof.set_Offset(curve, -1300);
                                Parameter BO = footprintRooff.LookupParameter("Base Offset From Level");
                                BO.SetValueString("22");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {

                }


                transaction.Commit();
            }

        }
        public void createfloor(Document doc, Element _l, IList<string[]> xYZ, string lvl)
        {
            CurveArray profile = new CurveArray();

            List<XYZ> collpnt = new List<XYZ>();
            List<string[]> levelpnt = new List<string[]>();
            Level level = (_l) as Level;

            foreach (string[] star in xYZ)
            {
                string strt = star[0];
                string end = star[1];
                string levl = star[2];
                Fwidth = star[3];
                FOffset = star[4];

                if (levl.ToString().Trim() == level.Name.Trim())
                {
                    myStrings = new[] { strt, end, levl, Fwidth, FOffset };
                    levelpnt.Add(myStrings);
                }
            }

            foreach (string[] star in levelpnt)
            {
                string strt = star[0];
                string end = star[1];
                levl = star[2];
                Fwidth = star[3];
                FOffset = star[4];

                if (levl.ToString().Trim() == level.Name.Trim())
                {

                    string Spoint = strt.Replace("(", "").Replace(")", "");
                    string Endpoint = end.Replace("(", "").Replace(")", "");

                    string[] start = Spoint.Split(',');
                    string[] en = Endpoint.Split(',');
                    //start point
                    double xa = Convert.ToDouble(start[0]);
                    double ya = Convert.ToDouble(start[1]);
                    double za = Convert.ToDouble(start[2]);

                    //end point 
                    double xb = Convert.ToDouble(en[0]);
                    double yb = Convert.ToDouble(en[1]);
                    double zb = Convert.ToDouble(en[2]);

                    XYZ point_a = new XYZ(xa, ya, za);
                    XYZ point_b = new XYZ(xb, yb, zb);

                    collpnt.Add(point_a);
                    collpnt.Add(point_b);
                }

            }
            // int n = collpnt.Count;

            for (int i = 0; i < collpnt.Count; i++)
            {
                Line line = Line.CreateBound(collpnt[i],
          collpnt[(i < collpnt.Count - 1) ? i + 1 : 0]);

                profile.Append(line);
            }

            XYZ normal = XYZ.BasisZ;
            //FloorType floorType = new FilteredElementCollector(doc).OfClass(typeof(FloorType)).First<Element>(e => e.Name.Equals("Generic - 12\"")) as FloorType;
            List<FloorType> floorTypes = new List<FloorType>();
            List<FloorType> GenericfloorTypes = new List<FloorType>();
            floorTypes = collfloorTypes(doc);

            foreach (FloorType wt in floorTypes) if (wt.Name.Contains("Generic")) GenericfloorTypes.Add(wt);
            using (Transaction transaction = new Transaction(doc, "Place Floor"))
            {
                if (!transaction.HasStarted()) transaction.Start();
                FailureHandlingOptions options = transaction.GetFailureHandlingOptions();
                MyPreProcessor preproccessor = new MyPreProcessor();
                options.SetClearAfterRollback(true);
                options.SetFailuresPreprocessor(preproccessor);
                transaction.SetFailureHandlingOptions(options);
                try
                {
                    if (profile.Size > 1)
                    {
                        IList<Parameter> Thick = new List<Parameter>();
                        foreach (FloorType ft in floorTypes)
                        {
                            bool presentcolumn = true;
                            foreach (FloorType ftype in floorTypes)
                            {
                                Parameter Ftypewidth = ftype.LookupParameter("Default Thickness");
                                Double fw = Ftypewidth.AsDouble();
                                _floorTypeWid = Math.Round((Double)fw, 4);

                                Double w = Convert.ToDouble(Fwidth);
                                _Wid = Math.Round((Double)w, 4);

                                if (_Wid == _floorTypeWid)
                                {
                                    Floor floorfmly = doc.Create.NewFloor(profile, ftype, level, true, normal);
                                    Parameter Heightoffset = floorfmly.LookupParameter("Height Offset From Level");
                                    Heightoffset.SetValueString(FOffset);
                                    presentcolumn = false;
                                }
                            }

                            if (presentcolumn)
                            {
                                if (_Wid != _floorTypeWid)
                                {

                                    foreach (FloorType wtyp in GenericfloorTypes)
                                    {
                                        FloorType newWallTyp = GenericfloorTypes[0].Duplicate("Custom FoorType" + Guid.NewGuid().ToString()) as FloorType;
                                        CompoundStructure cs = newWallTyp.GetCompoundStructure();
                                        int i = cs.GetFirstCoreLayerIndex();
                                        double thickness_to_set = Convert.ToDouble(Fwidth);
                                        cs.SetLayerWidth(i, thickness_to_set);
                                        newWallTyp.SetCompoundStructure(cs);

                                        Floor floorfmly = doc.Create.NewFloor(profile, newWallTyp, level, true, normal);
                                        Parameter Heightoffset = floorfmly.LookupParameter("Height Offset From Level");
                                        Heightoffset.SetValueString(FOffset);
                                        floorTypes.Add(newWallTyp);
                                        break;

                                    }

                                }
                            }
                            goto skip;
                        }
                    skip:;
                        transaction.Commit();
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("CapInfo", ex.Message);
                }

            }
        }
        public void createBeam(IList<string[]> collbeam, Element _l, Document doc)
        {
            List<XYZ> collpnt = new List<XYZ>();
            List<string[]> beampnt = new List<string[]>();
            Level level = (_l) as Level;
            foreach (string[] star in collbeam)
            {
                string Bstrt = star[0];
                string Bend = star[1];
                string Blevl = star[2];
                string Bwidth = star[3];
                string Bhight = star[4];

                if (Blevl.ToString().Trim() == level.Name.Trim())
                {
                    myStrings = new[] { Bstrt, Bend, Blevl, Bwidth, Bhight };
                    beampnt.Add(myStrings);
                }
            }
            foreach (string[] star in beampnt)
            {
                string Bstrt = star[0];
                string Bend = star[1];
                string Blevl = star[2];
                string Bwidth = star[3];
                string Bhight = star[4];

                double _w = Convert.ToDouble(Bwidth);
                double _h = Convert.ToDouble(Bhight);

                if (Blevl.ToString().Trim() == level.Name.Trim())
                {

                    string Spoint = Bstrt.Replace("(", "").Replace(")", "");
                    string Endpoint = Bend.Replace("(", "").Replace(")", "");
                    string[] start = Spoint.Split(',');
                    string[] end = Endpoint.Split(',');
                    //start point
                    double xa = Convert.ToDouble(start[0]);
                    double ya = Convert.ToDouble(start[1]);
                    double za = Convert.ToDouble(start[2]);

                    //end point 
                    double xb = Convert.ToDouble(end[0]);
                    double yb = Convert.ToDouble(end[1]);
                    double zb = Convert.ToDouble(end[2]);

                    XYZ point_a = new XYZ(xa, ya, za);
                    XYZ point_b = new XYZ(xb, yb, zb);



                    //Level level = doc.GetElement(view.LevelId) as Level;            // get a family symbol
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    collector.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralFraming);
                    bool presentcolumn = true;

                    //FamilySymbol gotSymbol = collector.FirstElement() as FamilySymbol;


                    Autodesk.Revit.DB.Curve beamLine = Line.CreateBound(point_a, point_b);

                    using (var transaction = new Transaction(doc))
                    {
                        transaction.Start("create Beam");

                        FailureHandlingOptions options = transaction.GetFailureHandlingOptions();
                        MyPreProcessor preproccessor = new MyPreProcessor();
                        options.SetClearAfterRollback(true);
                        options.SetFailuresPreprocessor(preproccessor);
                        transaction.SetFailureHandlingOptions(options);

                        foreach (FamilySymbol type in collector)
                        {
                            FamilySymbol gotSymbol = type as FamilySymbol;

                            try
                            {
                                Element et = doc.GetElement(type.Id);
                                Parameter h = et.LookupParameter("h");
                                if (h == null) continue;
                                Parameter w = et.LookupParameter("b");

                                double hh = h.AsDouble();
                                double ww = w.AsDouble();

                                if (_h == hh && _w == ww)
                                {
                                    type.Activate();

                                    //FamilyInstance instance = null;
                                    if (null != type)
                                    {

                                        if (!gotSymbol.IsActive)
                                        {
                                            gotSymbol.Activate();
                                            doc.Regenerate();
                                        }
                                        Level beamlevel = (level) as Level;
                                        instance = doc.Create.NewFamilyInstance(beamLine, gotSymbol, beamlevel, StructuralType.Beam);
                                        //Parameter zoffset = instance.LookupParameter("z Offset Value");
                                        //zoffset.Set(0);
                                        presentcolumn = false;
                                    }
                                }

                            }
                            catch (Exception ex)
                            {
                                TaskDialog.Show("CapInfo", ex.Message);
                            }
                        }


                        if (presentcolumn)
                        {
                            //string newFamilyName = "NEW REC TYPE" + " " + n;
                            //string newFamilyName = fm.ToString();
                            string newFamilyName = Guid.NewGuid().ToString();

                            try
                            {
                                FamilySymbol Btype = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralFraming).Cast<FamilySymbol>().Where(x => x.FamilyName.Contains("Beam")).FirstOrDefault(); // family type M_Concrete-Rectangular-Column
                                FamilySymbol beamType = Btype.Duplicate(newFamilyName) as FamilySymbol;
                                FamilySymbol gotSymbol = beamType as FamilySymbol;
                                Element e1 = doc.GetElement(beamType.Id);

                                Parameter h1 = e1.LookupParameter("h");
                                if (h1 != null)
                                {
                                    h1.Set(_h);
                                }
                                Parameter w1 = e1.LookupParameter("b");
                                if (w1 != null)
                                {
                                    w1.Set(_w);
                                }

                                beamType.Activate();

                                if (null != beamType)
                                {
                                    if (!gotSymbol.IsActive)
                                    {
                                        gotSymbol.Activate();
                                        doc.Regenerate();
                                    }
                                    Level beamlevel = (level) as Level;
                                    instance = doc.Create.NewFamilyInstance(beamLine, gotSymbol, beamlevel, StructuralType.Beam);

                                }
                            }
                            catch (Exception ex)
                            {
                                TaskDialog.Show("CapInfo", ex.Message + ex.StackTrace);
                            }

                        }



                        //if (!gotSymbol.IsActive)
                        //{
                        //    gotSymbol.Activate();
                        //    doc.Regenerate();
                        //}
                        //FamilyInstance instance = doc.Create.NewFamilyInstance(beamLine, gotSymbol, level, StructuralType.Beam);
                        //Parameter zoffset = instance.LookupParameter("z Offset Value");
                        //zoffset.Set(0);

                        transaction.Commit();
                    }
                }

                Parameter RLevel = instance.LookupParameter("Reference Level");
                string RefLevel = RLevel.AsValueString();

                using (Transaction transaction = new Transaction(doc))
                {
                    if (transaction.Start("Create") == TransactionStatus.Started)
                    {
                        FilteredElementCollector COLLFloor = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType();
                        foreach (Floor floorname in COLLFloor)
                        {
                            Floor flr = (Floor)floorname;

                            Parameter width = floorname.get_Parameter(BuiltInParameter.STRUCTURAL_FLOOR_CORE_THICKNESS);
                            string thickness = width.AsValueString();

                            string lvl = flr.LookupParameter("Level").AsValueString();
                            string HightOffset = flr.LookupParameter("Height Offset From Level").AsValueString();

                            Double Hgtoffset = Convert.ToDouble(HightOffset);
                            Double thick = Convert.ToDouble(thickness);

                            Double diffoffset = Hgtoffset - thick;

                            if (lvl.ToUpper() == RefLevel.ToString().ToUpper())
                            {

                                if (Hgtoffset == 0)
                                {
                                    Parameter Startoffset = instance.LookupParameter("Start Level Offset");
                                    Startoffset.Set(diffoffset);
                                    Parameter Endoffset = instance.LookupParameter("End Level Offset");
                                    Endoffset.Set(diffoffset);
                                }
                                else
                                {
                                    Parameter Startoffset = instance.LookupParameter("Start Level Offset");
                                    Startoffset.Set(diffoffset);
                                    Parameter Endoffset = instance.LookupParameter("End Level Offset");
                                    Endoffset.Set(diffoffset);
                                }


                            }
                        }
                        transaction.Commit();


                    }
                }
            }

        }
        private void createfillregion(Document doc, IList<string[]> point)
        {
            List<String> levlname = new List<String>();
            var colllevel = new FilteredElementCollector(doc).OfClass(typeof(Level)).OfCategory(BuiltInCategory.OST_Levels);
            foreach (string[] star in point)
            {
                string strt = star[0];

                string Spoint = strt.Replace("(", "").Replace(")", "");
                string[] start = Spoint.Split(',');

                //start point
                if (start.Count() != 3)
                {

                    foreach (Level levelname in colllevel)
                    {
                        if (levelname.Name.ToUpper().Trim() == strt.ToString().ToUpper().Trim())
                        {
                            _l = levelname;
                        }
                    }

                }

            }


            FilteredElementCollector fillRegionTypes = new FilteredElementCollector(doc).OfClass(typeof(FilledRegionType));

            Autodesk.Revit.DB.View viewL1 = null;
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewPlan));
            List<ViewPlan> ViewPlans = collector.Cast<ViewPlan>().ToList();
            //foreach (var lvlname in colllevel)
            //{
            using (Transaction t = new Transaction(doc, "Change View Model"))
            {
                t.Start();

                foreach (ViewPlan view in ViewPlans)
                {
                    Autodesk.Revit.DB.View v = view as Autodesk.Revit.DB.View;
                    if (v.Name.ToUpper() == _l.Name.ToUpper().ToString())
                    {
                        viewL1 = v;
                        break;
                    }
                }
                t.Commit();
            }


            foreach (FilledRegionType frt in fillRegionTypes)
            {
                List<CurveLoop> profileloops
                  = new List<CurveLoop>();

                List<XYZ> location = new List<XYZ>();

                XYZ point_a = null;
                foreach (string[] pt in point)
                {
                    string strt = pt[0];

                    string Spoint = strt.Replace("(", "").Replace(")", "");
                    string[] start = Spoint.Split(',');

                    //start point
                    if (start.Count() == 3)
                    {
                        double xa = Convert.ToDouble(start[0]);
                        double ya = Convert.ToDouble(start[1]);
                        double za = Convert.ToDouble(start[2]);

                        point_a = new XYZ(xa, ya, za);
                        location.Add(point_a);
                    }

                }

                CurveLoop profileloop = new CurveLoop();

                if (location.Count == 4)
                {
                    for (int i = 0; i < location.Count; i++)
                    {

                        Line line = Line.CreateBound(location[i],
              location[(i < 4 - 1) ? i + 1 : 0]);

                        profileloop.Append(line);
                    }
                    profileloops.Add(profileloop);

                    ElementId activeViewId = doc.ActiveView.Id;
                    Level level = (_l) as Level;
                    using (Transaction t = new Transaction(doc, "Place Column"))
                    {
                        t.Start();
                        FilledRegion filledRegion = FilledRegion.Create(doc, frt.Id, viewL1.Id, profileloops);
                        t.Commit();
                    }

                    break;
                }
            }
        }
        public void createCirclefillregion(Document doc, IList<string[]> point)
        {
            List<String> levlname = new List<String>();
            var colllevel = new FilteredElementCollector(doc).OfClass(typeof(Level)).OfCategory(BuiltInCategory.OST_Levels);


            foreach (Level lv in colllevel)
            {
                foreach (string[] star in point)
                {
                    string strt = star[0];
                    string radi = star[1];
                    radius = Convert.ToDouble(radi);

                    string levl = star[2];

                    if (levl.ToString().Trim() == lv.Name.Trim())
                    {
                        _l = lv;

                        string Spoint = strt.Replace("(", "").Replace(")", "");
                        string[] start = Spoint.Split(',');

                        //start point
                        if (start.Count() == 3)
                        {
                            double xa = Convert.ToDouble(start[0]);
                            double ya = Convert.ToDouble(start[1]);
                            double za = Convert.ToDouble(start[2]);

                            center = new XYZ(xa, ya, za);

                        }

                    }
                }
            }

            FilteredElementCollector fillRegionTypes = new FilteredElementCollector(doc).OfClass(typeof(FilledRegionType));

            Autodesk.Revit.DB.View viewL1 = null;
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewPlan));
            List<ViewPlan> ViewPlans = collector.Cast<ViewPlan>().ToList();
            //foreach (var lvlname in colllevel)
            //{
            using (Transaction t = new Transaction(doc, "Change View Model"))
            {
                t.Start();

                foreach (ViewPlan view in ViewPlans)
                {
                    Autodesk.Revit.DB.View v = view as Autodesk.Revit.DB.View;
                    if (v.Name.ToUpper() == _l.Name.ToUpper().ToString())
                    {
                        viewL1 = v;
                        break;
                    }
                }

                t.Commit();
            }

            foreach (FilledRegionType frt in fillRegionTypes)
            {
                //double startAngle = 0;
                //double endAngle = 360;
                XYZ xAxis = new XYZ(1, 0, 0);
                XYZ yAxis = new XYZ(0, 1, 0);

                double startAngle = 0;
                double endAngle = Math.PI;

                Arc a1 = Arc.Create(center, radius, startAngle, endAngle, xAxis, yAxis);
                Arc a2 = Arc.Create(center, radius, -endAngle, startAngle, xAxis, yAxis);

                //var circle = Arc.Create(center, radius, startAngle, endAngle, xAxis, yAxis);
                //profileloop.Append(circle);

                List<CurveLoop> profileloops = new List<CurveLoop>();
                CurveLoop profileloop = new CurveLoop();

                profileloop.Append(a1);
                profileloop.Append(a2);


                profileloops.Add(profileloop);


                using (Transaction transaction = new Transaction(doc))
                {
                    if (transaction.Start("Create filled region") == TransactionStatus.Started)
                    {
                        FilledRegion filledRegion1 = FilledRegion.Create(doc, frt.Id, viewL1.Id, profileloops);
                        transaction.Commit();
                    }
                }
                break;
            }
        }
        private void CreateWindow(Document doc, string fsFamilyName, string fsName, string levelName, string wposx, string wposy, IList<Element> walls)
        {
            FamilySymbol familySymbol = (from fs in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                                         where (fs.Family.Name == fsFamilyName && fs.Name == fsName)
                                         select fs).First();

            Level level = (from lvl in new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>() where (lvl.Name == levelName) select lvl).First();
            double x = double.Parse(wposx);
            double y = double.Parse(wposy);
            XYZ xyz = new XYZ(x, y, level.Elevation); FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Wall)); Wall wall = null;


            XYZ midpoint = null;
            foreach (Wall w in collector)
            {
                double proximity = (w.Location as LocationCurve).Curve.Distance(xyz);

                LocationCurve lc = w.Location as LocationCurve;
                Line line = lc.Curve as Line;
                XYZ p = line.GetEndPoint(0);
                XYZ q = line.GetEndPoint(1);
                XYZ v = q - p;
                midpoint = p + 0.5 * v;
                Parameter area = w.LookupParameter("Area");
                double _area = area.AsDouble();
                var _finalarea = string.Format("{0:0.00}", _area);
                double finalarea = Convert.ToDouble(_finalarea);
                if (_area > 290)
                {
                    wall = w; using (Transaction t = new Transaction(doc, "Create Window"))
                    {
                        t.Start();
                        if (!familySymbol.IsActive)
                        {
                            familySymbol.Activate();
                            doc.Regenerate();
                        }
                        FamilyInstance door = doc.Create.NewFamilyInstance(midpoint, familySymbol, wall, StructuralType.NonStructural);

                        t.Commit();

                    }
                }

            }
        }
        private void createDoor(IList<string[]> colldoor, Element _l, Document doc)
        {
            //FamilySymbol familySymbol = (from fs in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>() select fs).First();
            FamilySymbol familySymbol = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_Doors).Cast<FamilySymbol>().First();

            FilteredElementCollector collectordoor = new FilteredElementCollector(doc);
            collectordoor.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_Doors);



            //FamilySymbol familySymbol = (from fs in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
            //                             where (fs.Family.Name == "Single-Flush" && fs.Name == "40\" x 90\"")
            //                             select fs).First();
            //FamilySymbol familySymbolsmall = (from fs in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
            //                                  where (fs.Family.Name == "Single-Flush" && fs.Name == "30\" x 80\"")
            //                                  select fs).First();

            //FamilySymbol familySymboldoor = (from fs in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
            //                                 where (fs.Family.Name == "Single-Flush" && fs.Name == "35\" x 70\"")
            //                                 select fs).First();

            Level level = (from lvl in new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>() where (lvl.Name == _l.Name) select lvl).First();

            List<XYZ> collpnt = new List<XYZ>();
            List<string[]> doorpnt = new List<string[]>();
            Level lvls = (_l) as Level;

            foreach (string[] star in colldoor)
            {
                string dx = star[0];
                string dy = star[1];
                string dz = star[2];
                string dlevel = star[3];
                string walllength = star[4];
                string dw = star[5];
                string dh = star[6];

                _wlnth = Convert.ToDouble(walllength);
                //_wlnth = Math.Truncate(_wllenth);

                if (dlevel.ToString().Trim() == lvls.Name.Trim())
                {
                    myStrings = new[] { dx, dy, dz, dlevel, walllength, dw, dh };
                    doorpnt.Add(myStrings);
                }

                double x = double.Parse(dx);
                double y = double.Parse(dy);
                doorxyz = new XYZ(x, y, lvls.Elevation);
                collpnt.Add(doorxyz);
            }

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Wall));
            Wall wall = null;
            double distance = double.MaxValue;
            XYZ midpoint = null;
            XYZ midpoints = null;
            XYZ midpointdoor = null;

            foreach (Wall w in collector)
            {
                using (var transaction = new Transaction(doc))
                {
                    transaction.Start("create door");
                    if (_l.Id == w.LevelId)
                    {
                        Parameter walllength = w.LookupParameter("Length");
                        double _wallgth = walllength.AsDouble();
                        var _finallen = string.Format("{0:0.000}", _wallgth);
                        double finallength = Convert.ToDouble(_finallen);

                        //if (finallength == _wlnth)
                        //{

                        foreach (XYZ dxyz in collpnt)
                        {
                            double proximity = (w.Location as LocationCurve).Curve.Distance(dxyz);
                            LocationCurve lc = w.Location as LocationCurve;
                            Line line = lc.Curve as Line;

                            XYZ p = line.GetEndPoint(0);
                            XYZ q = line.GetEndPoint(1);
                            XYZ v = q - p;

                            midpoints = p + 0.5 * v;
                            midpoint = p + 0.3 * v;
                            midpointdoor = p + 0.35 * v;

                            Double x = midpoint.X;
                            Double y = midpoint.X;
                            Double z = midpoint.X;

                            Parameter area = w.LookupParameter("Area");


                            //double _area = area.AsDouble();
                            //var _finalarea = string.Format("{0:0.00}", _area);
                            //double finalarea = Convert.ToDouble(_finalarea);


                            //if (finalarea == 294.94 || finalarea == 45.63 || finalarea == 96.75 || finalarea == 768.58)
                            //{
                            foreach (string[] star in doorpnt)
                            {
                                string doorx = star[0];
                                string doory = star[1];
                                string doorz = star[2];
                                string doorlevel = star[3];
                                string walllenth = star[4];
                                string doorw = star[5];
                                string doorh = star[6];

                                double filength = Convert.ToDouble(walllenth);

                                if (filength == finallength)
                                {
                                    Double _dwth = Convert.ToDouble(doorw);
                                    _dw = Math.Truncate(_dwth);
                                    Double _dhght = Convert.ToDouble(doorh);
                                    _dh = Math.Truncate(_dhght);

                                    wall = w;
                                    bool presentdoor = true;


                                    foreach (FamilySymbol type in collectordoor)
                                    {
                                        FamilySymbol gotSymbol = type as FamilySymbol;

                                        try
                                        {
                                            Element et = doc.GetElement(type.Id);

                                            Parameter drh = et.LookupParameter("Height");
                                            if (drh == null) continue;
                                            Parameter drw = et.LookupParameter("Width");

                                            double hh = drh.AsDouble();
                                            double ww = drw.AsDouble();

                                            if (_dh == hh && _dw == ww)
                                            {
                                                type.Activate();


                                                if (null != type)
                                                {

                                                    if (!gotSymbol.IsActive)
                                                    {
                                                        gotSymbol.Activate();
                                                        doc.Regenerate();
                                                    }
                                                    FamilyInstance door = doc.Create.NewFamilyInstance(midpointdoor, gotSymbol, wall, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                    presentdoor = false;
                                                    goto Skip;
                                                }
                                            }

                                        }
                                        catch (Exception ex)
                                        {
                                            TaskDialog.Show("CapInfo", ex.Message);
                                        }
                                    }


                                    if (presentdoor)
                                    {
                                        string newFamilyName = Guid.NewGuid().ToString();
                                        try
                                        {
                                            FamilySymbol drtype = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_Doors).Cast<FamilySymbol>().Where(dx => dx.FamilyName.Contains("Flush")).FirstOrDefault(); // family type M_Concrete-Rectangular-Column
                                            FamilySymbol doorType = drtype.Duplicate(newFamilyName) as FamilySymbol;
                                            FamilySymbol gotSymbol = doorType as FamilySymbol;
                                            Element e1 = doc.GetElement(doorType.Id);

                                            Parameter h1 = e1.LookupParameter("Height");
                                            if (h1 != null)
                                            {
                                                h1.Set(_dh);
                                            }
                                            Parameter w1 = e1.LookupParameter("Width");
                                            if (w1 != null)
                                            {
                                                w1.Set(_dw);
                                            }

                                            doorType.Activate();


                                            if (null != doorType)
                                            {
                                                if (!gotSymbol.IsActive)
                                                {
                                                    gotSymbol.Activate();
                                                    doc.Regenerate();
                                                }
                                                FamilyInstance door = doc.Create.NewFamilyInstance(midpointdoor, gotSymbol, wall, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                goto Skip;
                                            }

                                        }

                                        catch (Exception ex)
                                        {
                                            TaskDialog.Show("CapInfo", ex.Message + ex.StackTrace);
                                        }

                                    }

                                }
                            }
                        }
                        //using (Transaction t = new Transaction(doc, "Create door"))
                        //{
                        //    t.Start();
                        //    if (!familySymbol.IsActive)
                        //    {
                        //        familySymbol.Activate();
                        //        doc.Regenerate();
                        //    }
                        //    FamilyInstance door = doc.Create.NewFamilyInstance(midpoint, familySymbol, wall, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                        //    t.Commit();
                        //}

                        //}
                        //if (finalarea == 45.63 || finalarea == 96.75)
                        //{
                        //    wall = w;
                        //    using (Transaction t = new Transaction(doc, "Create door"))
                        //    {
                        //        t.Start();
                        //        if (!familySymbolsmall.IsActive)
                        //        {
                        //            familySymbolsmall.Activate();
                        //            doc.Regenerate();
                        //        }
                        //        FamilyInstance door = doc.Create.NewFamilyInstance(midpoints, familySymbolsmall, wall, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                        //        t.Commit();
                        //    }
                        //}
                        //if (finalarea == 768.58)
                        //{
                        //    wall = w;
                        //    using (Transaction t = new Transaction(doc, "Create door"))
                        //    {
                        //        t.Start();
                        //        if (!familySymboldoor.IsActive)
                        //        {
                        //            familySymboldoor.Activate();
                        //            doc.Regenerate();
                        //        }
                        //        FamilyInstance door = doc.Create.NewFamilyInstance(midpointdoor, familySymboldoor, wall, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                        //        t.Commit();
                        //    }
                        //}
                    }
                //}
                Skip:;
                    transaction.Commit();
                }
            }

        }

        public void createwall(IList<string[]> wallcoll, string startpoint, string endpoint, Document doc, Element _l, string Level)
        {
            //var level_id = new ElementId(1526);
            Level level = (_l) as Level;
            // create line

            List<String> duplicatLevel = new List<String>();
            List<String> duplicat = new List<String>();

            List<WallType> oWallTypes = new List<WallType>();
            List<WallType> newWallType = new List<WallType>();
            List<WallType> GenericWallTypes = new List<WallType>();
            oWallTypes = GetWallTypes(doc);
            foreach (WallType wt in oWallTypes) if (wt.Name.Contains("Generic")) GenericWallTypes.Add(wt);
            //foreach (WallType wt in oWallTypes) if (wt.Name.Contains("Generic") || wt.Name.Contains("Curtain")) GenericWallTypes.Add(wt);
            var colllevel = new FilteredElementCollector(doc).OfClass(typeof(Level)).OfCategory(BuiltInCategory.OST_Levels);

            Level levelBottom = null;
            Level levelMiddle = null;
            Level levelTop = null;
            Level toplevel = null;



            foreach (string[] wallpoint in wallcoll)
            {
                string w_startpoint = wallpoint[0];
                string w_endpoint = wallpoint[1];
                levl = wallpoint[2];
                string height = wallpoint[3];

                string _Width = wallpoint[4];
                // string CurWall = wallpoint[5];

                if (levl.ToString() == level.Name)
                {

                    string Spoint = w_startpoint.Replace("(", "").Replace(")", "");
                    string Endpoint = w_endpoint.Replace("(", "").Replace(")", "");
                    string[] start = Spoint.Split(',');
                    string[] end = Endpoint.Split(',');
                    //start point
                    double xa = Convert.ToDouble(start[0]);
                    double ya = Convert.ToDouble(start[1]);
                    double za = Convert.ToDouble(start[2]);

                    //end point 
                    double xb = Convert.ToDouble(end[0]);
                    double yb = Convert.ToDouble(end[1]);
                    double zb = Convert.ToDouble(end[2]);

                    XYZ point_a = new XYZ(xa, ya, za);
                    XYZ point_b = new XYZ(xb, yb, zb);
                    // create line
                    Line line = Line.CreateBound(point_a, point_b);
                    var levelid = new FilteredElementCollector(doc).OfClass(typeof(Level)).OfCategory(BuiltInCategory.OST_Levels).Where(e => e.Name.ToUpper() == level.Name.ToUpper()).FirstOrDefault();
                    Wall wall;


                    using (var transaction = new Transaction(doc))
                    {
                        transaction.Start("create walls");
                        FailureHandlingOptions options = transaction.GetFailureHandlingOptions();
                        MyPreProcessor preproccessor = new MyPreProcessor();
                        options.SetClearAfterRollback(true);
                        options.SetFailuresPreprocessor(preproccessor);
                        transaction.SetFailureHandlingOptions(options);

                        wall = Wall.Create(doc, line, levelid.Id, false);
                        Parameter _area = wall.LookupParameter("Unconnected Height");
                        _area.SetValueString(height);
                        #region command

                        //Parameter _area = wall.LookupParameter("Unconnected Height");
                        //_area.SetValueString(height);


                        //if (height.ToString() == "3.88779527559055")
                        //{
                        //    Parameter BO = wall.LookupParameter("Base Offset");
                        //    BO.SetValueString("12");
                        //}
                        //if (height.ToString() == "3.72375328083989")
                        //{
                        //    Parameter BO = wall.LookupParameter("Base Offset");
                        //    BO.SetValueString("12.175");
                        //}
                        //if (height.ToString() == "2.46062992125984")
                        //{
                        //    Parameter BO = wall.LookupParameter("Base Offset");
                        //    BO.SetValueString("13.500");
                        //}

                        //if (height.ToString() == "22.8943569553806" && _Width == "0.820209973753281")
                        //{
                        //    Parameter BO = wall.LookupParameter("Base Offset");
                        //    BO.SetValueString("-6.993");
                        //}
                        //if (height.ToString() == "4.42913385826771")
                        //{
                        //    Parameter BO = wall.LookupParameter("Base Offset");
                        //    BO.SetValueString("11.580");
                        //}
                        //if (height.ToString() == "3.93700787401575")
                        //{
                        //    Parameter BO = wall.LookupParameter("Base Offset");
                        //    BO.SetValueString("-3.700");
                        //}
                        //if (height.ToString() == "18.0000000000000")
                        //{
                        //    Parameter BO = wall.LookupParameter("Base Offset");
                        //    BO.SetValueString("-2");
                        //}
                        //if (height.ToString() == "11.00656167979")
                        //{
                        //    Parameter BO = wall.LookupParameter("Base Offset");
                        //    BO.SetValueString("-2");
                        //}
                        //if (height.ToString() == "11.0065616797892")
                        //{
                        //    Parameter BO = wall.LookupParameter("Base Offset");
                        //    BO.SetValueString("-2");
                        //}
                        //if (height.ToString() == "4.47944281223773")
                        //{
                        //    Parameter BO = wall.LookupParameter("Base Offset");
                        //    BO.SetValueString("7.170");
                        //}



                        #endregion command

                        Element et = wall.Document.GetElement(wall.GetTypeId());

                        Wall wallwidth = ((Wall)wall);

                        //Parameter RLevel = wall.LookupParameter("Base Constraint");
                        //string RefLevel = RLevel.AsValueString();
                        //Parameter topconstraint = wall.LookupParameter("Top Constraint");

                        //FilteredElementCollector COLLFloor = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType();

                        //List<Level> topLevelname = colllevel.Cast<Level>().ToList<Level>();
                        //for (int i = 0; i < topLevelname.Count - 1; i++)
                        //{
                        //    if (RefLevel.ToString().Contains(topLevelname[i].Name))
                        //    {
                        //        var Levl = from vlv in colllevel where vlv.Name.Contains(topLevelname[i + 1].Name) select vlv;
                        //        List<Level> topLevel = Levl.Cast<Level>().ToList<Level>();

                        //        if (topLevel != null)

                        //            topconstraint.Set(topLevel[0].Id);

                        //        var flr = from fr in COLLFloor where fr.LookupParameter("Level").AsValueString().Equals(topLevelname[i + 1].Name) select fr;
                        //        List<Floor> fltlvl = flr.Cast<Floor>().ToList<Floor>();

                        //        Double width = fltlvl[0].get_Parameter(BuiltInParameter.STRUCTURAL_FLOOR_CORE_THICKNESS).AsDouble();

                        //        string HightOffset = fltlvl[0].LookupParameter("Height Offset From Level").AsValueString();

                        //        Double Hgtoffset = Convert.ToDouble(HightOffset);
                        //        Double thick = Convert.ToDouble(width);

                        //        Double diffoffset = Hgtoffset - thick;
                        //        Parameter topoffset = wall.LookupParameter("Top Offset");
                        //        topoffset.Set(diffoffset);

                        //    }
                        //}

                        foreach (WallType wt in oWallTypes)
                        {
                            //WallType wallType1 = doc.GetElement(wt.GetTypeId()) as WallType;
                            double nativeWitdh = wt.Width;
                            //double milimeterWidth = UnitUtils.ConvertFromInternalUnits(nativeWitdh, DisplayUnitType.DUT_MILLIMETERS);
                            //double milimeterWidth = UnitUtils.ConvertToInternalUnits(nativeWitdh, UnitTypeId.Millimeters);
                            double milimeterWidth = UnitUtils.ConvertFromInternalUnits(nativeWitdh, UnitTypeId.Millimeters);

                            bool presentcolumn = true;
                            foreach (WallType wt1 in oWallTypes)
                            {

                                w = Convert.ToDouble(_Width);
                                _Wid = Math.Round((Double)w, 4);

                                _nativeWitdth = wt1.Width;
                                Double _wallTypeWid = Math.Round((Double)_nativeWitdth, 4);

                                if (_Wid == _wallTypeWid)
                                {
                                    wall.ChangeTypeId(wt1.Id);
                                    presentcolumn = false;
                                }
                            }

                            if (presentcolumn)
                            {

                                if (_Wid != milimeterWidth)
                                {

                                    foreach (WallType wtyp in GenericWallTypes)
                                    {

                                        WallType newWallTyp = GenericWallTypes[0].Duplicate("Custom WallType" + Guid.NewGuid().ToString()) as WallType;
                                        CompoundStructure cs = newWallTyp.GetCompoundStructure();
                                        int i = cs.GetFirstCoreLayerIndex();

                                        double thickness_to_set = Convert.ToDouble(_Width);

                                        cs.SetLayerWidth(i, thickness_to_set);

                                        newWallTyp.SetCompoundStructure(cs);

                                        wall.ChangeTypeId(newWallTyp.Id);
                                        oWallTypes.Add(newWallTyp);
                                        break;
                                    }
                                }

                            }
                            goto skip;
                        }
                    skip:;

                        transaction.Commit();

                    }
                }
            }
        }
        private void createCwall(IList<string[]> wallcoll, string startpoint, string endpoint, Document doc, Element _l, string Level)
        {
            //var level_id = new ElementId(1526);
            Level level = (_l) as Level;
            // create line

            FilteredElementCollector Wallcollector = new FilteredElementCollector(doc);
            Wallcollector.OfClass(typeof(Wall));

            List<WallType> oWallTypes = new List<WallType>();
            List<WallType> newWallType = new List<WallType>();
            List<WallType> GenericWallTypes = new List<WallType>();
            oWallTypes = GetWallTypes(doc);
            foreach (WallType wt in oWallTypes) if (wt.FamilyName.Contains("Curtain")) GenericWallTypes.Add(wt);
            //foreach (WallType wt in oWallTypes) if (wt.Name.Contains("Generic") || wt.Name.Contains("Curtain")) GenericWallTypes.Add(wt);

            foreach (string[] wallpoint in wallcoll)
            {
                string w_startpoint = wallpoint[0];
                string w_endpoint = wallpoint[1];
                levl = wallpoint[2];
                string height = wallpoint[3];
                string _Width = wallpoint[4];


                if (levl.ToString() == level.Name)
                {
                    string Spoint = w_startpoint.Replace("(", "").Replace(")", "");
                    string Endpoint = w_endpoint.Replace("(", "").Replace(")", "");
                    string[] start = Spoint.Split(',');
                    string[] end = Endpoint.Split(',');
                    //start point
                    double xa = Convert.ToDouble(start[0]);
                    double ya = Convert.ToDouble(start[1]);
                    double za = Convert.ToDouble(start[2]);

                    //end point 
                    double xb = Convert.ToDouble(end[0]);
                    double yb = Convert.ToDouble(end[1]);
                    double zb = Convert.ToDouble(end[2]);

                    XYZ point_a = new XYZ(xa, ya, za);
                    XYZ point_b = new XYZ(xb, yb, zb);
                    // create line
                    Line line = Line.CreateBound(point_a, point_b);
                    var levelid = new FilteredElementCollector(doc).OfClass(typeof(Level)).OfCategory(BuiltInCategory.OST_Levels).Where(e => e.Name.ToUpper() == level.Name.ToUpper()).FirstOrDefault();
                    Wall wall;


                    using (var transaction = new Transaction(doc))
                    {
                        transaction.Start("create walls");
                        FailureHandlingOptions options = transaction.GetFailureHandlingOptions();
                        MyPreProcessor preproccessor = new MyPreProcessor();
                        options.SetClearAfterRollback(true);
                        options.SetFailuresPreprocessor(preproccessor);
                        transaction.SetFailureHandlingOptions(options);

                        //List<WallType> newWallType = new List<WallType>();
                        //List<WallType> GenericWallTypes = new List<WallType>();
                        //oWallTypes = GetWallTypes(doc);
                        //newWallTypes(doc, _Width, oWallTypes, transaction);

                        //wall = Wall.Create(doc, line, levelid.Id, false);

                        //Element et = wall.Document.GetElement(wall.GetTypeId());

                        //Wall wallwidth = ((Wall)wall);

                        //Parameter RLevel = wall.LookupParameter("Base Constraint");
                        //string RefLevel = RLevel.AsValueString();
                        //FilteredElementCollector COLLFloor = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType();


                        foreach (Wall wt in Wallcollector)
                        {

                            //WallType wallType1 = doc.GetElement(wt.GetTypeId()) as WallType;
                            double nativeWitdh = wt.Width;
                            //double milimeterWidth = UnitUtils.ConvertFromInternalUnits(nativeWitdh, DisplayUnitType.DUT_MILLIMETERS);

                            //bool presentcolumn = true;
                            //foreach (WallType wt1 in oWallTypes)
                            //{
                            Double w = Convert.ToDouble(_Width);
                            _Wid = Math.Round((Double)w, 4);

                            _nativeWitdth = wt.Width;
                            Double _wallTypeWid = Math.Round((Double)_nativeWitdth, 4);

                            if (_Wid == _wallTypeWid)
                            {
                                foreach (WallType wtyp in GenericWallTypes)
                                {
                                    if (wtyp.Name == "Exterior Glazing")
                                    {
                                        wt.ChangeTypeId(wtyp.Id);

                                        Element et = wt.Document.GetElement(wt.GetTypeId());
                                        Parameter RLevell = et.LookupParameter("Automatically Embed");
                                        RLevell.Set(1);
                                        break;
                                    }
                                }
                                goto skip;

                            }



                            //if (presentcolumn)
                            //{

                            //    if (_Wid != milimeterWidth)
                            //    {

                            //        foreach (WallType wtyp in GenericWallTypes)
                            //        {
                            //            WallType newWallTyp = GenericWallTypes[0].Duplicate("Custom WallType" + Guid.NewGuid().ToString()) as WallType;
                            //            CompoundStructure cs = newWallTyp.GetCompoundStructure();
                            //            int i = cs.GetFirstCoreLayerIndex();
                            //            double thickness_to_set = Convert.ToDouble(_Width);
                            //            cs.SetLayerWidth(i, thickness_to_set);
                            //            newWallTyp.SetCompoundStructure(cs);

                            //            wall.ChangeTypeId(newWallTyp.Id);
                            //            oWallTypes.Add(newWallTyp);
                            //            break;

                            //        }

                            //    }

                            //}
                            //goto skip;


                        }

                    skip:;

                        transaction.Commit();
                    }






                }



            }
        }
        private void createColumn(Document doc, Element _l)
        {
            try
            {
                string path = Path.Combine(Path.GetFullPath(System.Reflection.Assembly.GetExecutingAssembly().Location).Remove(Path.GetFullPath(System.Reflection.Assembly.GetExecutingAssembly().Location).Count() - Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location).Count()), @"Template\ColumnFamilySchdule.xlsx");
                DataTable dataTable = xlsxToDT(path);

                List<ElementId> idsExclude = new List<ElementId>();

                FilteredElementCollector collectorUsed = new FilteredElementCollector(doc);
                //FilteredElementCollector fillRegions = new FilteredElementCollector(doc, doc.ActiveView.Id).OfClass(typeof(FilledRegion));
                FilteredElementCollector fillRegions = new FilteredElementCollector(doc).OfClass(typeof(FilledRegion));
                //FilteredElementCollector coltxtt = collectorUsed.OfClass(typeof(TextElement)).OfCategory(BuiltInCategory.OST_TextNotes);
                FilteredElementCollector coltxt = new FilteredElementCollector(doc).OfClass(typeof(TextElement)).OfCategory(BuiltInCategory.OST_TextNotes);
                FilteredElementCollector colview = collectorUsed.OfClass(typeof(ViewPlan));

                var colllevel = new FilteredElementCollector(doc).OfClass(typeof(Level)).OfCategory(BuiltInCategory.OST_Levels);

                Level level = (_l) as Level;
                Autodesk.Revit.DB.View viewL1 = null;
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfClass(typeof(ViewPlan));
                List<ViewPlan> ViewPlans = collector.Cast<ViewPlan>().ToList();
                IList<ElementId> collFilledRegion = new List<ElementId>();

                //foreach (var lvlname in colllevel)
                //{
                using (Transaction t = new Transaction(doc, "Change View Model"))
                {
                    t.Start();

                    foreach (ViewPlan view in ViewPlans)
                    {
                        //if (view.ViewType == ViewType.pl)
                        //{
                        Autodesk.Revit.DB.View v = view as Autodesk.Revit.DB.View;
                        if (v.Name.ToUpper() == level.Name.ToUpper().ToString())
                        {
                            viewL1 = v;
                            levelfillRegions = new FilteredElementCollector(doc, viewL1.Id).OfClass(typeof(FilledRegion));
                            break;
                        }
                        //}
                    }

                    t.Commit();
                }

                IList<int> lstelementids = new List<int>();
                foreach (Element tname in coltxt)
                {
                    string txtname = (((Autodesk.Revit.DB.TextElement)tname).Text).Trim();

                    foreach (DataRow row in dataTable.Rows)
                    {
                        string columntext = row[0].ToString();

                        if (columntext.ToUpper().Trim() == txtname.ToUpper().Trim())
                        {
                            using (Transaction transaction = new Transaction(doc, "Place Column"))
                            {
                                if (!transaction.HasStarted()) transaction.Start();
                                try
                                {

                                    BoundingBoxXYZ tbb = tname.get_BoundingBox(viewL1);
                                    XYZ tbbmidSum = tbb.Max + tbb.Min;
                                    XYZ tbbmid = new XYZ(tbbmidSum.X / 2, tbbmidSum.Y / 2, 0);
                                    double arcradius = .8;
                                    double startAngle = 0;
                                    double endAngle = 360;
                                    XYZ xAxis = new XYZ(1, 0, 0);
                                    XYZ yAxis = new XYZ(0, 1, 0);

                                    Frame frame = new Frame();
                                    Plane geomPlane = Plane.Create(frame);
                                    Arc txtarc = Arc.Create(tbbmid, arcradius, startAngle, endAngle, xAxis, yAxis);

                                    SketchPlane sketch = SketchPlane.Create(doc, geomPlane);
                                    ModelCurve acircle = doc.Create.NewModelCurve(txtarc, sketch) as ModelCurve;
                                    idsExclude.Add(acircle.Id);

                                    foreach (FilledRegion _filledRegion in levelfillRegions)
                                    {
                                        collFilledRegion.Add(_filledRegion.Id);
                                        FilledRegion fill = _filledRegion as FilledRegion;
                                        BoundingBoxXYZ bb = _filledRegion.get_BoundingBox(viewL1);
                                        XYZ bbmidSum = bb.Max + bb.Min;
                                        XYZ bbmid = new XYZ(bbmidSum.X / 2, bbmidSum.Y / 2, 0);
                                        Outline outline = new Outline(bb.Min, bb.Max);
                                        ElementQuickFilter fbb = new BoundingBoxIntersectsFilter(outline);

                                        IList<CurveLoop> interloops = fill.GetBoundaries();
                                        bool isfound = false;
                                        foreach (CurveLoop floop in interloops)
                                        {
                                            foreach (Curve fcurve in floop)
                                            {
                                                SetComparisonResult set = txtarc.Intersect(fcurve);

                                                if (set.ToString().ToUpper() == "Overlap".ToUpper())
                                                {
                                                    IList<double> arc = new List<double>();
                                                    IList<double> p1 = new List<double>();
                                                    IList<CurveLoop> loops = fill.GetBoundaries();
                                                    //schfamilytype = row[1].ToString();
                                                    foreach (CurveLoop loop in loops)
                                                    {

                                                        foreach (Curve curve in loop)
                                                        {
                                                            if (curve is Autodesk.Revit.DB.Line)
                                                            {
                                                                double p2 = curve.Length;
                                                                p1.Add(p2);
                                                            }

                                                            if (curve is Autodesk.Revit.DB.Arc)
                                                            {
                                                                double radius = 0;
                                                                radius = ((Autodesk.Revit.DB.Arc)curve).Radius;

                                                                if (radius != 0)
                                                                    arc.Add(radius);
                                                            }
                                                        }
                                                    }

                                                    double _w = 0;
                                                    double _h = 0;
                                                    double arcc = 0;

                                                    if (p1.Count() == 4)
                                                    {
                                                        _w = p1.Min();
                                                        _h = p1.Max();
                                                    }
                                                    else if (arc.Count() == 2)
                                                    {
                                                        arcc = arc.Max();
                                                    }

                                                    BoundingBoxXYZ _bound = _filledRegion.get_BoundingBox(viewL1);
                                                    XYZ midSum = _bound.Max + _bound.Min;
                                                    XYZ mid = new XYZ(midSum.X / 2, midSum.Y / 2, 0);


                                                    if (p1.Count() == 4)
                                                    {
                                                        //PlaceRecColumne(doc, mid, _h, _w);
                                                        PlaceRecColumne(doc, mid, _h, _w, schfamilytype, level);
                                                    }
                                                    else if (arc.Count() == 2)
                                                    {
                                                        PlaceCircleColumn(doc, mid, arcc, schfamilytype, level);
                                                    }
                                                    lstelementids.Add(_filledRegion.Id.IntegerValue);
                                                    isfound = true;
                                                    goto skip; ;
                                                }
                                            }
                                        }
                                        //skip:;

                                    }
                                skip:;
                                    transaction.Commit();
                                }
                                catch (Exception ex)
                                {
                                    TaskDialog.Show("CapInfo", ex.Message);
                                }
                            }


                        }
                    }
                    //break;
                }
                using (Transaction transaction = new Transaction(doc, "Place Column"))
                {
                    if (!transaction.HasStarted()) transaction.Start();
                    try
                    {
                        foreach (FilledRegion _filledRegion in levelfillRegions)
                        {
                            collFilledRegion.Add(_filledRegion.Id);

                            FilledRegion fill = _filledRegion as FilledRegion;
                            if (!lstelementids.Contains(_filledRegion.Id.IntegerValue))
                            {
                                IList<double> arc = new List<double>();
                                IList<double> p1 = new List<double>();
                                IList<CurveLoop> loops = fill.GetBoundaries();
                                List<XYZ> lstPts = new List<XYZ>();
                                //string schfamilytype = row[1].ToString();
                                foreach (CurveLoop loop in loops)
                                {
                                    foreach (Curve curve in loop)
                                    {
                                        lstPts.AddRange(curve.Tessellate());

                                        if (curve is Autodesk.Revit.DB.Line)
                                        {
                                            double p2 = curve.Length;
                                            p1.Add(p2);
                                        }

                                        if (curve is Autodesk.Revit.DB.Arc)
                                        {
                                            double radius = 0;
                                            radius = ((Autodesk.Revit.DB.Arc)curve).Radius;
                                            if (radius != 0)
                                                arc.Add(radius);
                                        }
                                    }
                                }

                                List<XYZ> tmp = lstPts.OrderBy(x => x.Z).ToList();
                                List<XYZ> tmp1 = lstPts.OrderBy(x => x.X).ToList();

                                double height = Math.Abs(tmp[0].Z - tmp[tmp.Count - 1].Z);
                                double width = Math.Abs(tmp1[0].X - tmp1[tmp1.Count - 1].X);

                                double _w = 0;
                                double _h = 0;
                                double arcc = 0;

                                if (p1.Count() == 4)
                                {
                                    _w = p1.Min();
                                    _h = p1.Max();

                                    if (_w != width)
                                    {
                                        _w = width;
                                        _h = p1.Min();
                                    }
                                    else
                                    {
                                        _w = p1.Min();
                                        _h = p1.Max();

                                    }

                                    //_w = p1.Min();
                                    //_h = p1.Max();

                                }
                                else if (arc.Count() == 2)
                                {
                                    arcc = arc.Max();
                                }

                                BoundingBoxXYZ _bound = _filledRegion.get_BoundingBox(viewL1);
                                XYZ midSum = _bound.Max + _bound.Min;
                                XYZ mid = new XYZ(midSum.X / 2, midSum.Y / 2, 0);

                                if (p1.Count() == 4)
                                {
                                    PlaceRecColumne(doc, mid, _w, _h, schfamilytype, level);
                                }
                                else if (arc.Count() == 2)
                                {
                                    PlaceCircleColumn(doc, mid, arcc, schfamilytype, level);
                                }
                                lstelementids.Add(_filledRegion.Id.IntegerValue);

                            }
                        }
                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("CapInfo", ex.Message);
                    }
                }

                using (Transaction transaction = new Transaction(doc, "Delete FilledRegion"))
                {
                    if (!transaction.HasStarted()) transaction.Start();
                    try
                    {
                        doc.Delete(collFilledRegion);
                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("CapInfo", ex.Message);

                    }
                }

            }

            catch (Exception ex) { TaskDialog.Show("", "" + ex.Message); }

        }
        public void PlaceRecColumne(Document doc, XYZ origine, Double _j, Double _k, string fm, Element lvlname)
        {

            Double _w = Math.Round((Double)_j, 2);
            Double _h = Math.Round((Double)_k, 2);

            //double _h = _j;
            //double _w = _k;

            bool presentcolumn = true;
            FilteredElementCollector collector1 = new FilteredElementCollector(doc);
            collector1.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralColumns);
            //FamilySymbol type = collector1.FirstElement() as FamilySymbol;

            foreach (FamilySymbol type in collector1)
            {

                try
                {
                    Element et = doc.GetElement(type.Id);

                    Parameter h = et.LookupParameter("h");
                    if (h == null) continue;
                    Parameter w = et.LookupParameter("b");

                    double hh = h.AsDouble();
                    double ww = w.AsDouble();

                    if (_h == hh && _w == ww)
                    {
                        type.Activate();
                        //var level = new FilteredElementCollector(doc).OfClass(typeof(Level)).OfCategory(BuiltInCategory.OST_Levels).Where(e => e.Name.ToUpper() == "Level 1".ToUpper()).FirstOrDefault();


                        FamilyInstance instance = null;
                        if (null != type)
                        {
                            Level _l = (lvlname) as Level;
                            instance = doc.Create.NewFamilyInstance(origine, type, _l, StructuralType.Column);
                            presentcolumn = false;
                        }
                    }

                }
                catch (Exception ex)
                {
                    TaskDialog.Show("CapInfo", ex.Message);
                }
            }


            if (presentcolumn)
            {

                //string newFamilyName = "NEW REC TYPE" + " " + n;
                //string newFamilyName = fm.ToString();
                string newFamilyName = Guid.NewGuid().ToString();

                try
                {
                    //FamilySymbol ctype = collector1.FirstElement() as FamilySymbol;
                    FamilySymbol ctype = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralColumns).Cast<FamilySymbol>().Where(x => x.FamilyName.Contains("Rectangular")).FirstOrDefault(); // family type M_Concrete-Rectangular-Column
                    if (ctype != null)
                    {
                        FamilySymbol columnType = ctype.Duplicate(newFamilyName) as FamilySymbol;

                        Element e1 = doc.GetElement(columnType.Id);

                        Parameter h1 = e1.LookupParameter("h");
                        if (h1 != null)
                        {
                            h1.Set(_h);
                        }
                        Parameter w1 = e1.LookupParameter("b");
                        if (w1 != null)
                        {
                            w1.Set(_w);
                        }

                        columnType.Activate();
                        //var level1 = new FilteredElementCollector(doc).OfClass(typeof(Level)).OfCategory(BuiltInCategory.OST_Levels).Where(e => e.Name.ToUpper() == "Level 1".ToUpper()).FirstOrDefault();
                        FamilyInstance instance1 = null;
                        if (null != columnType)
                        {

                            Level _l = (lvlname) as Level;
                            instance1 = doc.Create.NewFamilyInstance(origine, columnType, _l, StructuralType.Column);
                        }
                    }
                    //n++;
                }

                catch (Exception ex)
                {
                    TaskDialog.Show("CapInfo", ex.Message + ex.StackTrace);
                }

            }

        }
        public void PlaceCircleColumn(Document doc, XYZ origine, Double s, string fm, Element lvlname)
        {
            //double _w = Math.Round(double.Parse(txtcirwidth.Text)) / 304.8;
            // double _w = double.Parse(txtcirwidth.Text) / 304.8;
            //double _w = s * 2;

            Double _a = s * 2;
            Double _arc = Math.Round((Double)_a, 1);

            bool presentcolumn = true;
            FilteredElementCollector collector1 = new FilteredElementCollector(doc);
            collector1.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralColumns);

            foreach (FamilySymbol atype in collector1)
            {
                try
                {
                    //FamilySymbol atype = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralColumns).Cast<FamilySymbol>().Where(x => x.FamilyName.Contains("Round")).FirstOrDefault();
                    Element et = doc.GetElement(atype.Id);

                    Parameter w = et.LookupParameter("b");
                    if (w == null) continue;
                    double ww = w.AsDouble();
                    if (_arc == ww)
                    {
                        atype.Activate();

                        FamilyInstance instance = null;
                        if (null != atype)
                        {
                            Level _l = (lvlname) as Level;
                            instance = doc.Create.NewFamilyInstance(origine, atype, _l, StructuralType.Column);
                            presentcolumn = false;
                        }
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("owbt", ex.Message);
                }
            }

            if (presentcolumn)
            {
                //string newFamilyName = fm.ToString();
                string newFamilyName = Guid.NewGuid().ToString();
                try
                {

                    // FamilySymbol ctype = collector1.FirstElement() as FamilySymbol;
                    FamilySymbol ctype = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralColumns).Cast<FamilySymbol>().Where(x => x.FamilyName.Contains("Round")).FirstOrDefault(); // family type

                    FamilySymbol columnType1 = ctype.Duplicate(newFamilyName) as FamilySymbol;

                    Element e1 = doc.GetElement(columnType1.Id);
                    Parameter w1 = e1.LookupParameter("b");
                    w1.Set(_arc);

                    columnType1.Activate();

                    // var level1 = new FilteredElementCollector(doc).OfClass(typeof(Level)).OfCategory(BuiltInCategory.OST_Levels).Where(e => e.Name.ToUpper() == "Level 1".ToUpper()).FirstOrDefault();
                    FamilyInstance instance1 = null;
                    if (null != columnType1)
                    {
                        Level _l = (lvlname) as Level;
                        instance1 = doc.Create.NewFamilyInstance(origine, columnType1, _l, StructuralType.Column);
                    }
                    //  a++;
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("", ex.Message);
                }
            }
        }
        public DataTable xlsxToDT(string path)
        {
            DataTable dt1 = new DataTable();
            try
            {
                using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    XSSFWorkbook hssfworkbook = new XSSFWorkbook(file);
                    ISheet sheet = hssfworkbook.GetSheetAt(0);
                    //  ISheet sheet = rdoElec.Checked ? hssfworkbook.GetSheet("Parcel_Elec") : hssfworkbook.GetSheet("Parcel_Mech");
                    //if (sheet == null)
                    //    sheet = hssfworkbook.GetSheetAt(0);

                    IRow headerRow = sheet.GetRow(0);
                    System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
                    int colCount = headerRow.LastCellNum;
                    int rowCount = sheet.LastRowNum;

                    for (int c = 0; c < colCount; c++)
                    {
                        try
                        {
                            dt1.Columns.Add(headerRow.GetCell(c).ToString());
                        }
                        catch (Exception ex) { dt1.Columns.Add(Guid.NewGuid().ToString()); }
                    }
                    bool skipReadingHeaderRow = rows.MoveNext();
                    while (rows.MoveNext())
                    {
                        {
                            IRow row = (XSSFRow)rows.Current;
                            DataRow dr = dt1.NewRow();
                            for (int i = 0; i < colCount; i++)
                            {

                                ICell cell = row.GetCell(i);

                                if (cell != null)
                                {
                                    if (cell.CellType == NPOI.SS.UserModel.CellType.Formula)
                                    {
                                        if (((NPOI.XSSF.UserModel.XSSFCell)cell).CachedFormulaResultType == NPOI.SS.UserModel.CellType.Numeric)
                                            dr[i] = cell.NumericCellValue.ToString();
                                        else if (((NPOI.XSSF.UserModel.XSSFCell)cell).CachedFormulaResultType == NPOI.SS.UserModel.CellType.String)
                                            dr[i] = cell.StringCellValue.ToString();
                                    }
                                    else
                                        dr[i] = cell.ToString();
                                }
                            }
                            dt1.Rows.Add(dr);
                        }
                    }
                    hssfworkbook = null;
                    sheet = null;
                }
            }
            catch (Exception ex) { System.Windows.Forms.MessageBox.Show(ex.Message.ToString()); }
            return dt1;
        }
        public static void deletelevel(Document document)
        {
            int deleted = 0;
            var colllevel = new FilteredElementCollector(document).OfClass(typeof(Level)).OfCategory(BuiltInCategory.OST_Levels);

            List<ElementId> elementsToBeDeleted = new List<ElementId>();
            using (Transaction transaction = new Transaction(document, "Place"))
            {
                if (!transaction.HasStarted()) transaction.Start();

                foreach (Element element in colllevel)
                {
                    elementsToBeDeleted.Add(element.Id);
                    deleted++;
                }
                document.Delete(elementsToBeDeleted);
                transaction.Commit();
            }
        }
        public static List<WallType> GetWallTypes(Autodesk.Revit.DB.Document doc)
        {
            using (Transaction transaction = new Transaction(doc, "WallType"))
            {
                if (!transaction.HasStarted()) transaction.Start();

                List<WallType> oWallTypes = new List<WallType>();
                try
                {
                    FilteredElementCollector collector
                        = new FilteredElementCollector(doc);

                    FilteredElementIterator itor = collector
                        .OfClass(typeof(HostObjAttributes))
                        .GetElementIterator();

                    // Reset the iterator
                    itor.Reset();

                    // Iterate through each family
                    while (itor.MoveNext())
                    {
                        Autodesk.Revit.DB.HostObjAttributes oSystemFamilies =
                        itor.Current as Autodesk.Revit.DB.HostObjAttributes;

                        if (oSystemFamilies == null) continue;

                        // Get the family's category
                        Category oCategory = oSystemFamilies.Category;

                        // Process if the category is found
                        if (oCategory != null)
                        {
                            if (oCategory.Name == "Walls")
                            {
                                WallType oWallType = oSystemFamilies as WallType;
                                if (oWallType != null) oWallTypes.Add(oWallType);
                            }
                        }
                    } //while itor.NextMove()

                    transaction.Commit();
                    return oWallTypes;
                }
                catch (Exception)
                {
                    //MessageBox.Show( ex.Message );
                    return oWallTypes = new List<WallType>();
                }
            }
        }
        public static List<FloorType> collfloorTypes(Autodesk.Revit.DB.Document doc)
        {
            using (Transaction transaction = new Transaction(doc, "floortype"))
            {
                if (!transaction.HasStarted()) transaction.Start();

                List<FloorType> ofloorTypes = new List<FloorType>();
                try
                {
                    FilteredElementCollector collector
                        = new FilteredElementCollector(doc);

                    FilteredElementIterator itor = collector
                        .OfClass(typeof(HostObjAttributes))
                        .GetElementIterator();

                    // Reset the iterator
                    itor.Reset();

                    // Iterate through each family
                    while (itor.MoveNext())
                    {
                        Autodesk.Revit.DB.HostObjAttributes oSystemFamilies =
                        itor.Current as Autodesk.Revit.DB.HostObjAttributes;

                        if (oSystemFamilies == null) continue;

                        // Get the family's category
                        Category oCategory = oSystemFamilies.Category;

                        // Process if the category is found
                        if (oCategory != null)
                        {
                            if (oCategory.Name == "Floors")
                            {
                                FloorType ofloorType = oSystemFamilies as FloorType;
                                if (ofloorType != null) ofloorTypes.Add(ofloorType);
                            }
                        }
                    } //while itor.NextMove()

                    transaction.Commit();
                    return ofloorTypes;
                }
                catch (Exception)
                {
                    //MessageBox.Show( ex.Message );
                    return ofloorTypes = new List<FloorType>();
                }
            }
        }
        private static void EncFile(string filename)
        {
            try
            {
                string text = File.ReadAllText(filename);
                string EncryptionKey = "CAPINFO";
                byte[] clearBytes = Encoding.Unicode.GetBytes(text);
                using (Aes encryptor = Aes.Create())
                {
                    Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] {
            0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76});
                    encryptor.Key = pdb.GetBytes(32);
                    encryptor.IV = pdb.GetBytes(16);
                    using (MemoryStream ms = new MemoryStream())
                    {
                        using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write))
                        {
                            cs.Write(clearBytes, 0, clearBytes.Length);
                            cs.Close();
                        }
                        text = Convert.ToBase64String(ms.ToArray());
                    }
                }
                using (StreamWriter writetext = new StreamWriter(filename))
                {
                    writetext.WriteLine(text);
                }
            }
            catch (Exception ex) { MessageBox.Show("ExportCap", "" + ex.Message + ex.StackTrace); }

        }
        private static void DecFile(string filename)
        {
            try
            {

                string text = File.ReadAllText(filename);
                string EncryptionKey = "CAPINFO";
                text = text.Replace(" ", "+");
                byte[] cipherBytes = Convert.FromBase64String(text);
                using (Aes encryptor = Aes.Create())
                {
                    Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] {
            0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76});
                    encryptor.Key = pdb.GetBytes(32);
                    encryptor.IV = pdb.GetBytes(16);
                    using (MemoryStream ms = new MemoryStream())
                    {
                        using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write))
                        {
                            cs.Write(cipherBytes, 0, cipherBytes.Length);
                            cs.Close();
                        }
                        text = Encoding.Unicode.GetString(ms.ToArray());


                        using (StreamWriter writetext = new StreamWriter(filename))
                        {
                            writetext.WriteLine(text);
                        }
                    }
                }

            }
            catch (Exception ex) { MessageBox.Show("ExportCap", "" + ex.Message + ex.StackTrace); }
        }
        public class MyPreProcessor : IFailuresPreprocessor
        {
            FailureProcessingResult IFailuresPreprocessor.PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                String transactionName = failuresAccessor.GetTransactionName();

                IList<FailureMessageAccessor> fmas = failuresAccessor.GetFailureMessages();

                if (fmas.Count == 0)
                    return FailureProcessingResult.Continue;

                if (transactionName.Equals("EXEMPLE"))
                {
                    foreach (FailureMessageAccessor fma in fmas)
                    {
                        if (fma.GetSeverity() == FailureSeverity.Error)
                        {
                            failuresAccessor.DeleteAllWarnings();
                            return FailureProcessingResult.ProceedWithRollBack;
                        }
                        else
                        {
                            failuresAccessor.DeleteWarning(fma);
                        }

                    }
                }
                else
                {
                    foreach (FailureMessageAccessor fma in fmas)
                    {
                        failuresAccessor.DeleteAllWarnings();
                    }
                }
                return FailureProcessingResult.Continue;
            }
        }
    }
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Newstartup : IExternalCommand
    {
        private string caplocation;
        public string[] myStrings;
        public string Thickness = null;
        public FilteredElementCollector levelfillRegions;
        public Level _l;
        private double _floorTypeWid;
        private double w;
        private double _Wid;
        private double _nativeWitdth;
        private FamilyInstance instance;
        private int levelvalue;
        private RoofType rooftype;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            UIApplication uiapp = commandData.Application;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            var path = doc.PathName;


            Autodesk.Revit.DB.Units units = doc.GetUnits();
            //FormatOptions fo = units.GetFormatOptions(UnitType.UT_Length); //you specify which unit you are interested in setting/reading
            //FormatOptions nFt = new FormatOptions();
            //fo.DisplayUnits = DisplayUnitType.DUT_CUBIC_FEET;
            //units.SetFormatOptions(UnitType.UT_Length, fo);
            //doc.SetUnits(units);
            Autodesk.Revit.DB.View view = doc.ActiveView;

           // deletelevel(doc);

            #region Collect elements

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
            IList<Element> Walls = collector.WherePasses(filter).WhereElementIsNotElementType().ToElements();
            ICollection<Element> doors = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).ToElements();
            ICollection<Element> Windows = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).ToElements();
            FilteredElementCollector collectorbeams = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType();
            FilteredElementCollector COLLSLAB = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType();
            FilteredElementCollector fillRegions = new FilteredElementCollector(doc).OfClass(typeof(FilledRegion));
            var colllevel = new FilteredElementCollector(doc).OfClass(typeof(Level)).OfCategory(BuiltInCategory.OST_Levels);
            IList<Element> Wallstype = collector.WherePasses(filter).ToElements();

            #endregion Collect elements

            #region open Dilog
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"C:\",
                Title = "Browse Cap Files",
                CheckFileExists = true,
                CheckPathExists = true,
                DefaultExt = "cap",
                Filter = "Cap Files (*.cap)|*.cap",
                FilterIndex = 2,
                RestoreDirectory = true,
                ReadOnlyChecked = true,
                ShowReadOnly = true
            };
            #endregion open Dilog

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                caplocation = openFileDialog1.FileName;

                StringBuilder result = new StringBuilder();

                //Load xml
                //XDocument xdoc = XDocument.Load(caplocation);
                XElement rootElement = XElement.Load(caplocation);


                #region Level
                if (rootElement.Element("CAP").Element("Levels") != null)
                {
                    //Run query
                    var lv1s = rootElement.Element("CAP").Element("Levels").Elements().ToList();

                    foreach (var lvlname in lv1s)
                    {
                        var lstle = lvlname.Element("Level_Value");
                        var lName = lvlname.Element("Level_Name").Value.ToString();
                        var lNo = lvlname.Element("Level_Number").Value.ToString();
                        double lvldepth = 0;
                        double.TryParse(lstle.Value.ToString(), out lvldepth);
                        using (Transaction transaction = new Transaction(doc, "LevelCreation"))
                        {
                            if (!transaction.HasStarted()) transaction.Start();
                            var lvlcollection = colllevel.Where(_layer => _layer.Name == "Level " + lNo).Select(i => i).ToList();
                            if (lvlcollection.Count > 0)
                            {
                                Level level = lvlcollection[0] as Level;
                                level.Elevation = lvldepth / 305;
                            }
                            else
                            {
                                Level _l = Level.Create(doc, lvldepth / 305);
                                _l.Name = "Level " + lNo;
                            }

                            //_l.Elevation = lvldepth/100;
                            transaction.Commit();
                        }

                    }

                    TaskDialog.Show("2D-3D", "Level Created");
                }
                #endregion Level

                #region Column
                if (rootElement.Element("CAP").Element("Columns") != null)
                {
                    //Run query
                    var collColumns = rootElement.Element("CAP").Element("Columns").Elements().ToList();


                    List<XYZ> collColumnscollpnt = new List<XYZ>();
                    List<string[]> Columnspnt = new List<string[]>();

                    levelvalue = 1;
                    foreach (var columnname in collColumns)
                    {
                        var lstle = columnname.Element("Level_Value");
                        var lName = columnname.Element("Level_Name").Value.ToString();
                        var lNo = columnname.Element("Level_Number").Value.ToString();
                        var cPoints = columnname.Elements("cPoint");

                        var lvlcollection = colllevel.Where(_layer => _layer.Name == "Level " + lNo).Select(i => i).ToList();
                        Level level = lvlcollection[0] as Level;

                        //int.TryParse(lName, out levelvalue);
                        //levelvalue;
                        var colllevels = new FilteredElementCollector(doc).OfClass(typeof(Level)).OfCategory(BuiltInCategory.OST_Levels);
                        var collvl = colllevels.Where(_layer => _layer.Name == "Level " + levelvalue).Select(i => i).ToList();

                        Level nxtlvl = collvl[0] as Level;
                        levelvalue++;
                        foreach (string cPoint in cPoints)
                        {
                            string point = cPoint.Replace("<cPoint>", "").Replace("</cPoint>", "");
                            myStrings = new[] { point };
                            Columnspnt.Add(myStrings);
                        }
                        using (Transaction transaction = new Transaction(doc, "Place Column"))
                        {
                            if (!transaction.HasStarted()) transaction.Start();
                            FailureHandlingOptions options = transaction.GetFailureHandlingOptions();
                            MyPreProcessor preproccessor = new MyPreProcessor();
                            options.SetClearAfterRollback(true);
                            options.SetFailuresPreprocessor(preproccessor);
                            transaction.SetFailureHandlingOptions(options);
                            try
                            {
                                foreach (string[] star in Columnspnt)
                                {
                                    string fltpnt = star[0].ToString();
                                    string[] point = fltpnt.Split(')');

                                    string mid = point[0].Replace("(", "");
                                    var height = point[1].Replace("(", "");
                                    var width = point[2].Replace("(", "");
                                    var family = point[3].Replace("(", "");

                                    double _w = Convert.ToDouble(width);
                                    double _h = Convert.ToDouble(height);

                                    string[] start = mid.Split(',');

                                    double xa = Convert.ToDouble(start[0]);
                                    double ya = Convert.ToDouble(start[1]);
                                    double za = Convert.ToDouble(start[2]);

                                    XYZ origine = new XYZ(xa, ya, za);

                                    bool presentcolumn = true;
                                    FilteredElementCollector collector1 = new FilteredElementCollector(doc);
                                    collector1.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralColumns);


                                    foreach (FamilySymbol type in collector1)
                                    {

                                        try
                                        {
                                            Element et = doc.GetElement(type.Id);

                                            Parameter h = et.LookupParameter("h");
                                            if (h == null) continue;
                                            Parameter w = et.LookupParameter("b");

                                            double hh = h.AsDouble();
                                            double ww = w.AsDouble();

                                            if (_h == hh && _w == ww)
                                            {
                                                type.Activate();
                                                //var level = new FilteredElementCollector(doc).OfClass(typeof(Level)).OfCategory(BuiltInCategory.OST_Levels).Where(e => e.Name.ToUpper() == "Level 1".ToUpper()).FirstOrDefault();


                                                FamilyInstance instance = null;
                                                if (null != type)
                                                {
                                                    Level _l = (level) as Level;
                                                    instance = doc.Create.NewFamilyInstance(origine, type, _l, StructuralType.Column);
                                                    var toplevel = instance.LookupParameter("Top Level");
                                                    toplevel.Set(nxtlvl.Id);
                                                    var topoffset = instance.LookupParameter("Top Offset");
                                                    topoffset.SetValueString("0");

                                                    presentcolumn = false;
                                                }
                                            }

                                        }
                                        catch (Exception ex)
                                        {
                                            TaskDialog.Show("CapInfo", ex.Message);
                                        }
                                    }


                                    if (presentcolumn)
                                    {

                                        string newFamilyName = Guid.NewGuid().ToString();

                                        try
                                        {
                                            FamilySymbol ctype = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralColumns).Cast<FamilySymbol>().Where(x => x.FamilyName == family).FirstOrDefault(); // family type M_Concrete-Rectangular-Column
                                            if (ctype != null)
                                            {
                                                FamilySymbol columnType = ctype.Duplicate(newFamilyName) as FamilySymbol;

                                                Element e1 = doc.GetElement(columnType.Id);

                                                Parameter h1 = e1.LookupParameter("h");
                                                if (h1 != null)
                                                {
                                                    h1.Set(_h);
                                                }
                                                Parameter w1 = e1.LookupParameter("b");
                                                if (w1 != null)
                                                {
                                                    w1.Set(_w);
                                                }

                                                columnType.Activate();
                                                //var level1 = new FilteredElementCollector(doc).OfClass(typeof(Level)).OfCategory(BuiltInCategory.OST_Levels).Where(e => e.Name.ToUpper() == "Level 1".ToUpper()).FirstOrDefault();
                                                FamilyInstance instance1 = null;
                                                if (null != columnType)
                                                {

                                                    Level _l = (level) as Level;
                                                    instance1 = doc.Create.NewFamilyInstance(origine, columnType, _l, StructuralType.Column);
                                                    var toplevel = instance1.LookupParameter("Top Level");
                                                    toplevel.Set(nxtlvl.Id);
                                                    var topoffset = instance1.LookupParameter("Top Offset");
                                                    topoffset.SetValueString("0");
                                                }
                                            }
                                            //n++;
                                        }

                                        catch (Exception ex)
                                        {
                                            TaskDialog.Show("CapInfo", ex.Message + ex.StackTrace);
                                        }

                                    }

                                }
                                transaction.Commit();
                            }
                            catch (Exception ex)
                            {

                            }
                        }

                    }

                    TaskDialog.Show("CapInfo", "Column Created");
                }

                #endregion Column

                #region Floor

                if (rootElement.Element("CAP").Element("Floors") != null)
                {
                    //Run query
                    var floors = rootElement.Element("CAP").Element("Floors").Elements().ToList();


                    List<XYZ> collpnts = new List<XYZ>();
                    List<string> collpntsting = new List<string>();
                    List<string[]> floorpnt = new List<string[]>();
                    CurveArray profile = new CurveArray();
                    foreach (var floorname in floors)
                    {
                        var lstle = floorname.Element("Level_Value");
                        var lName = floorname.Element("Level_Name").Value.ToString();
                        var lNo = floorname.Element("Level_Number").Value.ToString();
                        var width = floorname.Element("Width").Value.ToString();
                        var offSet = floorname.Element("Offset").Value.ToString();
                        var famy = floorname.Element("familyname").Value.ToString();
                        var typ = floorname.Element("type").Value.ToString();

                        string family = famy.Replace("(", "").Replace(")", "");
                        string type = typ.Replace("(", "").Replace(")", "");


                        var cPoints = floorname.Elements("cPoint");

                        var lvlcollection = colllevel.Where(_layer => _layer.Name == "Level " + lNo).Select(i => i).ToList();
                        Level level = lvlcollection[0] as Level;

                        foreach (string cPoint in cPoints)
                        {
                            string point = cPoint.Replace("<cPoint>", "").Replace("</cPoint>", "");
                            myStrings = new[] { point };
                            floorpnt.Add(myStrings);
                        }
                        foreach (string[] star in floorpnt)
                        {

                            string roof = star[0].ToString();
                            string[] point = roof.Split(')');

                            foreach (string pnt in point)
                            {
                                string points = pnt.Replace("(", "");
                                collpntsting.Add(points);

                            }
                            foreach (string pnt in collpntsting)
                            {
                                string[] p1 = pnt.Split(',');
                                if (p1[0] != "")
                                {
                                    double a1 = Convert.ToDouble(p1[0]);
                                    double b1 = Convert.ToDouble(p1[1]);
                                    double c1 = Convert.ToDouble(p1[2]);
                                    XYZ ptt1 = new XYZ(a1, b1, c1);

                                    collpnts.Add(ptt1);
                                }
                            }

                        }

                        for (int i = 0; i < collpnts.Count; i++)
                        {
                            Line line = Line.CreateBound(collpnts[i],
                      collpnts[(i < collpnts.Count - 1) ? i + 1 : 0]);

                            profile.Append(line);
                        }
                        collpnts.Clear();
                        collpntsting.Clear();
                        floorpnt.Clear();

                        XYZ normal = XYZ.BasisZ;
                        // FloorType floorType = new FilteredElementCollector(doc).OfClass(typeof(FloorType)).First<Element>(e => e.Name.Equals("Generic - 12\"")) as FloorType;
                        List<FloorType> floorTypes = new List<FloorType>();
                        List<FloorType> GenericfloorTypes = new List<FloorType>();
                        floorTypes = collfloorTypes(doc);

                        //foreach (FloorType wt in floorTypes) if (wt.Name.Contains("Generic")) GenericfloorTypes.Add(wt);
                        foreach (FloorType wt in floorTypes) if (wt.FamilyName == family ) GenericfloorTypes.Add(wt);
                        using (Transaction transaction = new Transaction(doc, "Place Floor"))
                        {
                            if (!transaction.HasStarted()) transaction.Start();
                            FailureHandlingOptions options = transaction.GetFailureHandlingOptions();
                            MyPreProcessor preproccessor = new MyPreProcessor();
                            options.SetClearAfterRollback(true);
                            options.SetFailuresPreprocessor(preproccessor);
                            transaction.SetFailureHandlingOptions(options);
                            try
                            {
                                if (profile.Size > 1)
                                {
                                    IList<Parameter> Thick = new List<Parameter>();
                                    foreach (FloorType ft in floorTypes)
                                    {
                                        bool presentcolumn = true;
                                        foreach (FloorType ftype in floorTypes)
                                        {
                                            Parameter Ftypewidth = ftype.LookupParameter("Default Thickness");
                                            Double fw = Ftypewidth.AsDouble();
                                            _floorTypeWid = Math.Round((Double)fw, 4);

                                            Double w = Convert.ToDouble(width);
                                            _Wid = Math.Round((Double)w, 4);

                                            if (_Wid == _floorTypeWid)
                                            {
                                                Floor floorfmly = doc.Create.NewFloor(profile, ftype, level, true, normal);
                                                Parameter Heightoffset = floorfmly.LookupParameter("Height Offset From Level");
                                                Heightoffset.SetValueString(offSet);
                                                presentcolumn = false;
                                            }
                                        }

                                        if (presentcolumn)
                                        {

                                            if (_Wid != _floorTypeWid)
                                            {

                                                foreach (FloorType wtyp in GenericfloorTypes)
                                                {
                                                    FloorType newWallTyp = GenericfloorTypes[0].Duplicate("Custom FoorType" + Guid.NewGuid().ToString()) as FloorType;
                                                    CompoundStructure cs = newWallTyp.GetCompoundStructure();
                                                    int i = cs.GetFirstCoreLayerIndex();
                                                    double thickness_to_set = Convert.ToDouble(width);
                                                    cs.SetLayerWidth(i, thickness_to_set);
                                                    newWallTyp.SetCompoundStructure(cs);

                                                    Floor floorfmly = doc.Create.NewFloor(profile, newWallTyp, level, true, normal);
                                                    Parameter Heightoffset = floorfmly.LookupParameter("Height Offset From Level");
                                                    Heightoffset.SetValueString(offSet);
                                                    floorTypes.Add(newWallTyp);
                                                    break;

                                                }

                                            }
                                        }
                                        goto skip;
                                    }
                                skip:;
                                    transaction.Commit();
                                    profile.Clear();
                                }

                            }
                            catch (Exception ex)
                            {
                                TaskDialog.Show("CapInfoc", ex.StackTrace.ToString());
                                TaskDialog.Show("CapInfo", ex.Message);
                            }

                        }
                    }

                    TaskDialog.Show("2D-3D", "Floor Created");
                }
                #endregion Floor

                #region Beams

                if (rootElement.Element("CAP").Element("Beams") != null)
                {
                    //Run query
                    var collBeams = rootElement.Element("CAP").Element("Beams").Elements().ToList();


                    List<XYZ> beamcollpnt = new List<XYZ>();
                    List<string[]> beampnt = new List<string[]>();
                    CurveArray beamprofile = new CurveArray();

                    foreach (var beamname in collBeams)
                    {
                        var lstle = beamname.Element("Level_Value");
                        var lName = beamname.Element("Level_Name").Value.ToString();
                        var lNo = beamname.Element("Level_Number").Value.ToString();
                        var cPoints = beamname.Elements("cPoint");

                        var lvlcollection = colllevel.Where(_layer => _layer.Name == "Level " + lNo).Select(i => i).ToList();
                        Level level = lvlcollection[0] as Level;

                        foreach (string cPoint in cPoints)
                        {
                            string point = cPoint.Replace("<cPoint>", "").Replace("</cPoint>", "");
                            myStrings = new[] { point };
                            beampnt.Add(myStrings);
                        }
                        foreach (string[] star in beampnt)
                        {
                            string fltpnt = star[0].ToString();
                            string[] point = fltpnt.Split(')');

                            string Spoint = point[0].Replace("(", "");
                            string Endpoint = point[1].Replace("(", "");
                            var width = point[2].Replace("(", "");
                            var height = point[3].Replace("(", "");
                            var family = point[4].Replace("(", "");


                            double _w = Convert.ToDouble(width);
                            double _h = Convert.ToDouble(height);


                            string[] start = Spoint.Split(',');
                            string[] en = Endpoint.Split(',');
                            double xa = Convert.ToDouble(start[0]);
                            double ya = Convert.ToDouble(start[1]);
                            double za = Convert.ToDouble(start[2]);


                            double xb = Convert.ToDouble(en[0]);
                            double yb = Convert.ToDouble(en[1]);
                            double zb = Convert.ToDouble(en[2]);
                            XYZ point_a = new XYZ(xa, ya, za);
                            XYZ point_b = new XYZ(xb, yb, zb);

                            Line line = Line.CreateBound(point_a, point_b);

                            FilteredElementCollector beamcollector = new FilteredElementCollector(doc);
                            beamcollector.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralFraming);
                            bool presentcolumn = true;

                            Autodesk.Revit.DB.Curve beamLine = Line.CreateBound(point_a, point_b);

                            using (var transaction = new Transaction(doc))
                            {
                                transaction.Start("create Beam");

                                FailureHandlingOptions options = transaction.GetFailureHandlingOptions();
                                MyPreProcessor preproccessor = new MyPreProcessor();
                                options.SetClearAfterRollback(true);
                                options.SetFailuresPreprocessor(preproccessor);
                                transaction.SetFailureHandlingOptions(options);

                                foreach (FamilySymbol type in collector)
                                {
                                    FamilySymbol gotSymbol = type as FamilySymbol;

                                    try
                                    {
                                        Element et = doc.GetElement(type.Id);
                                        Parameter h = et.LookupParameter("h");
                                        if (h == null) continue;
                                        Parameter w = et.LookupParameter("b");

                                        double hh = h.AsDouble();
                                        double ww = w.AsDouble();

                                        if (_h == hh && _w == ww)
                                        {
                                            type.Activate();

                                            //FamilyInstance instance = null;
                                            if (null != type)
                                            {

                                                if (!gotSymbol.IsActive)
                                                {
                                                    gotSymbol.Activate();
                                                    doc.Regenerate();
                                                }
                                                Level beamlevel = (level) as Level;
                                                instance = doc.Create.NewFamilyInstance(beamLine, gotSymbol, beamlevel, StructuralType.Beam);
                                                presentcolumn = false;
                                            }
                                        }

                                    }
                                    catch (Exception ex)
                                    {
                                        TaskDialog.Show("CapInfo", ex.Message);
                                    }
                                }


                                if (presentcolumn)
                                {
                                    string newFamilyName = Guid.NewGuid().ToString();

                                    try
                                    {
                                        FamilySymbol Btype = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralFraming).Cast<FamilySymbol>().Where(x => x.FamilyName == family).FirstOrDefault(); // family type M_Concrete-Rectangular-Column
                                        FamilySymbol beamType = Btype.Duplicate(newFamilyName) as FamilySymbol;
                                        FamilySymbol gotSymbol = beamType as FamilySymbol;
                                        Element e1 = doc.GetElement(beamType.Id);

                                        Parameter h1 = e1.LookupParameter("h");
                                        if (h1 != null)
                                        {
                                            h1.Set(_h);
                                        }
                                        Parameter w1 = e1.LookupParameter("b");
                                        if (w1 != null)
                                        {
                                            w1.Set(_w);
                                        }

                                        beamType.Activate();

                                        if (null != beamType)
                                        {
                                            if (!gotSymbol.IsActive)
                                            {
                                                gotSymbol.Activate();
                                                doc.Regenerate();
                                            }
                                            Level beamlevel = (level) as Level;
                                            instance = doc.Create.NewFamilyInstance(beamLine, gotSymbol, beamlevel, StructuralType.Beam);

                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        TaskDialog.Show("CapInfo", ex.Message + ex.StackTrace);
                                    }
                                }
                                transaction.Commit();
                            }

                        }

                        Parameter RLevel = instance.LookupParameter("Reference Level");
                        string RefLevel = RLevel.AsValueString();

                        using (Transaction transaction = new Transaction(doc))
                        {
                            if (transaction.Start("Create") == TransactionStatus.Started)
                            {
                                FilteredElementCollector COLLFloor = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType();
                                foreach (Floor floorname in COLLFloor)
                                {
                                    Floor flr = (Floor)floorname;

                                    Parameter width = floorname.get_Parameter(BuiltInParameter.STRUCTURAL_FLOOR_CORE_THICKNESS);
                                    //string thickness = width.AsValueString();
                                    double thickness = width.AsDouble();
                                    //double feet = Convert.ToDouble(thickness);


                                    string lvl = flr.LookupParameter("Level").AsValueString();
                                    // string HightOffset = flr.LookupParameter("Height Offset From Level").AsValueString();
                                    string HightOffset = "0";

                                    Double Hgtoffset = Convert.ToDouble(HightOffset);
                                    Double thick = Convert.ToDouble(thickness);

                                    Double diffoffset = Hgtoffset - thick;

                                    if (lvl.ToUpper() == RefLevel.ToString().ToUpper())
                                    {

                                        if (Hgtoffset == 0)
                                        {
                                            Parameter Startoffset = instance.LookupParameter("Start Level Offset");
                                            Startoffset.Set(diffoffset);
                                            Parameter Endoffset = instance.LookupParameter("End Level Offset");
                                            Endoffset.Set(diffoffset);
                                        }
                                        else
                                        {
                                            Parameter Startoffset = instance.LookupParameter("Start Level Offset");
                                            Startoffset.Set(diffoffset);
                                            Parameter Endoffset = instance.LookupParameter("End Level Offset");
                                            Endoffset.Set(diffoffset);
                                        }


                                    }
                                }
                                transaction.Commit();


                            }
                        }
                    }

                    TaskDialog.Show("2D-3D", "Beam Created");
                }
                #endregion Beams

                #region Wall

                if (rootElement.Element("CAP").Element("Walls") != null)
                {
                    //Run query
                    var collWalls = rootElement.Element("CAP").Element("Walls").Elements().ToList();



                    List<XYZ> wallcollpnt = new List<XYZ>();
                    List<string[]> wallpnt = new List<string[]>();
                    CurveArray wallprofile = new CurveArray();

                    List<WallType> oWallTypes = new List<WallType>();
                    List<WallType> newWallType = new List<WallType>();
                    List<WallType> GenericWallTypes = new List<WallType>();
                    oWallTypes = GetWallTypes(doc);
                    //foreach (WallType wt in oWallTypes) if (wt.Name.Contains("Generic")) GenericWallTypes.Add(wt);

                    foreach (var wallname in collWalls)
                    {
                        var lstle = wallname.Element("Level_Value");
                        var lName = wallname.Element("Level_Name").Value.ToString();
                        var lNo = wallname.Element("Level_Number").Value.ToString();
                        var cPoints = wallname.Elements("cPoint");

                        var lvlcollection = colllevel.Where(_layer => _layer.Name == "Level " + lNo).Select(i => i).ToList();
                        Level level = lvlcollection[0] as Level;

                        foreach (string cPoint in cPoints)
                        {
                            string point = cPoint.Replace("<cPoint>", "").Replace("</cPoint>", "");
                            myStrings = new[] { point };
                            wallpnt.Add(myStrings);
                        }
                        foreach (string[] star in wallpnt)
                        {
                            string fltpnt = star[0].ToString();
                            string[] point = fltpnt.Split(')');

                            string Spoint = point[0].Replace("(", "");
                            string Endpoint = point[1].Replace("(", ""); ;
                            var height = point[2].Replace("(", "");
                            var width = point[3].Replace("(", "");
                            var walltype = point[4].Replace("(", "");


                            foreach (WallType wt in oWallTypes) if (wt.Name.Contains(walltype)) GenericWallTypes.Add(wt);

                            string[] start = Spoint.Split(',');
                            string[] en = Endpoint.Split(',');
                            double xa = Convert.ToDouble(start[0]);
                            double ya = Convert.ToDouble(start[1]);
                            double za = Convert.ToDouble(start[2]);


                            double xb = Convert.ToDouble(en[0]);
                            double yb = Convert.ToDouble(en[1]);
                            double zb = Convert.ToDouble(en[2]);
                            XYZ point_a = new XYZ(xa, ya, za);
                            XYZ point_b = new XYZ(xb, yb, zb);

                            Line line = Line.CreateBound(point_a, point_b);
                            Wall wall;


                            using (var transaction = new Transaction(doc))
                            {
                                transaction.Start("create walls");
                                FailureHandlingOptions options = transaction.GetFailureHandlingOptions();
                                MyPreProcessor preproccessor = new MyPreProcessor();
                                options.SetClearAfterRollback(true);
                                options.SetFailuresPreprocessor(preproccessor);
                                transaction.SetFailureHandlingOptions(options);

                                wall = Wall.Create(doc, line, level.Id, false);

                                Parameter _area = wall.LookupParameter("Unconnected Height");
                                _area.SetValueString(height);



                                Element et = wall.Document.GetElement(wall.GetTypeId());

                                Wall wallwidth = ((Wall)wall);


                                foreach (WallType wt in oWallTypes)
                                {

                                    double nativeWitdh = wt.Width;
                                    double milimeterWidth = UnitUtils.ConvertFromInternalUnits(nativeWitdh, UnitTypeId.Millimeters);

                                    bool presentcolumn = true;
                                    foreach (WallType wt1 in oWallTypes)
                                    {

                                        w = Convert.ToDouble(width);
                                        _Wid = Math.Round((Double)w, 4);

                                        _nativeWitdth = wt1.Width;
                                        Double _wallTypeWid = Math.Round((Double)_nativeWitdth, 4);

                                        if (_Wid == _wallTypeWid)
                                        {
                                            wall.ChangeTypeId(wt1.Id);
                                            presentcolumn = false;
                                        }
                                    }

                                    if (presentcolumn)
                                    {

                                        if (_Wid != milimeterWidth)
                                        {

                                            //foreach (WallType wtyp in GenericWallTypes)
                                            //{

                                            //    WallType newWallTyp = GenericWallTypes[0].Duplicate("Custom WallType" + Guid.NewGuid().ToString()) as WallType;
                                            //    CompoundStructure cs = newWallTyp.GetCompoundStructure();
                                            //    int i = cs.GetFirstCoreLayerIndex();

                                            //    double thickness_to_set = Convert.ToDouble(width);

                                            //    cs.SetLayerWidth(i, thickness_to_set);

                                            //    newWallTyp.SetCompoundStructure(cs);

                                            //    wall.ChangeTypeId(newWallTyp.Id);
                                            //    oWallTypes.Add(newWallTyp);
                                            //    break;
                                            //}
                                        }

                                    }
                                    goto skip;
                                }
                            skip:;

                                doc.Regenerate();
                                doc.AutoJoinElements();
                                transaction.Commit();

                            }
                        }
                        wallpnt.Clear();
                    }

                    TaskDialog.Show("2D-3D", "Wall Created");
                }
                #endregion Wall

                #region Door

                if (rootElement.Element("CAP").Element("Doors") != null)
                {

                    //var coldoor = rootElement.Element("CAP").Element("Doors").Elements().ToList();




                    //List<string[]> doorspnt = new List<string[]>();

                    //levelvalue = 1;
                    //foreach (var doorname in coldoor)
                    //{
                    //    var lstle = doorname.Element("Level_Value");
                    //    var lName = doorname.Element("Level_Name").Value.ToString();
                    //    var lNo = doorname.Element("Level_Number").Value.ToString();
                    //    var cPoints = doorname.Elements("cPoint");

                    //    var lvlcollection = colllevel.Where(_layer => _layer.Name == "Level " + lNo).Select(i => i).ToList();
                    //    Level level = lvlcollection[0] as Level;

                    //    Level levelname = (from lvl in new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>() where (lvl.Name == level.Name) select lvl).First();

                    //    foreach (string cPoint in cPoints)
                    //    {
                    //        string point = cPoint.Replace("<cPoint>", "").Replace("</cPoint>", "");
                    //        myStrings = new[] { point };
                    //        doorspnt.Add(myStrings);
                    //    }

                    //    foreach (string[] star in doorspnt)
                    //    {
                    //        string fltpnt = star[0].ToString();
                    //        string[] point = fltpnt.Split(')');

                    //        string location = point[0].Replace("(", "");
                    //        var family = point[1].Replace("(", "");
                    //        var type = point[2].Replace("(", "");

                    //        FamilySymbol familySymbol = (from fs in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                    //                                     where (fs.Family.Name == family && fs.Name == type)
                    //                                     select fs).First();


                    //        FilteredElementCollector wallcollector = new FilteredElementCollector(doc);
                    //        wallcollector.OfClass(typeof(Wall)); Wall wall = null;

                    //        XYZ midpoint = null;

                    //        foreach (Wall w in collector)
                    //        {
                    //            LocationCurve lc = w.Location as LocationCurve;
                    //            Line line = lc.Curve as Line;
                    //            XYZ p = line.GetEndPoint(0);
                    //            XYZ q = line.GetEndPoint(1);
                    //            XYZ v = q - p;
                    //            midpoint = p + 0.5 * v;
                    //            Parameter area = w.LookupParameter("Length");
                    //            String length = area.AsValueString();
                    //            //Parameter area = w.LookupParameter("Area");
                    //            //double _area = area.AsDouble();
                    //            //var _finalarea = string.Format("{0:0.00}", _area);
                    //            //double finalarea = Convert.ToDouble(_finalarea);
                    //            //if (finalarea == 45.06 || finalarea == 188.25)

                    //            if (length == "26")
                    //            {
                    //                wall = w;
                    //                using (Transaction t = new Transaction(doc, "Create Door"))
                    //                {
                    //                    t.Start();
                    //                    if (!familySymbol.IsActive)
                    //                    {
                    //                        familySymbol.Activate();
                    //                        doc.Regenerate();
                    //                    }
                    //                    FamilyInstance door = doc.Create.NewFamilyInstance(midpoint, familySymbol, wall, StructuralType.NonStructural);
                    //                    t.Commit();
                    //                }

                    //            }

                    //        }
                    //    }
                    //}
                    //TaskDialog.Show("2D-3D", "Door Created");
                }

                #endregion Door

                #region Window
                if (rootElement.Element("CAP").Element("Windows") != null)
                {
                    var colWindows = rootElement.Element("CAP").Element("Windows").Elements().ToList();




                    List<string[]> Windowspnt = new List<string[]>();

                    levelvalue = 1;
                    foreach (var Windowsname in colWindows)
                    {
                        var lstle = Windowsname.Element("Level_Value");
                        var lName = Windowsname.Element("Level_Name").Value.ToString();
                        var lNo = Windowsname.Element("Level_Number").Value.ToString();
                        var cPoints = Windowsname.Elements("cPoint");

                        var lvlcollection = colllevel.Where(_layer => _layer.Name == "Level " + lNo).Select(i => i).ToList();
                        Level level = lvlcollection[0] as Level;

                        foreach (string cPoint in cPoints)
                        {
                            string point = cPoint.Replace("<cPoint>", "").Replace("</cPoint>", "");
                            myStrings = new[] { point };
                            Windowspnt.Add(myStrings);
                        }

                        foreach (string[] star in Windowspnt)
                        {
                            string fltpnt = star[0].ToString();
                            string[] point = fltpnt.Split(')');

                            string location = point[0].Replace("(", "");
                            var family = point[1].Replace("(", "");
                            var type = point[2].Replace("(", "");
                            var offset = point[3].Replace("(", "");

                            FamilySymbol familySymbol = (from fs in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                                                         where (fs.Family.Name == family && fs.Name == type)
                                                         select fs).First();

                            FilteredElementCollector wallcollector = new FilteredElementCollector(doc);
                            wallcollector.OfClass(typeof(Wall)); Wall wall = null;

                            XYZ midpoint = null;

                            foreach (Wall w in wallcollector)
                            {

                                if (w.LevelId == level.Id)
                                {
                                    LocationCurve lc = w.Location as LocationCurve;
                                    Line line = lc.Curve as Line;
                                    XYZ p = line.GetEndPoint(0);
                                    XYZ q = line.GetEndPoint(1);
                                    XYZ v = q - p;
                                    midpoint = p + 0.5 * v;

                                    Parameter area = w.LookupParameter("Length");
                                    String length = area.AsValueString();

                                    //Parameter area = w.LookupParameter("Area");
                                    //double _area = area.AsDouble();
                                    //var _finalarea = string.Format("{0:0.00}", _area);
                                    //double finalarea = Convert.ToDouble(_finalarea);


                                    XYZ orgin = new XYZ(37.183281470, -0.470712638, -292.939016393);
                                    if (length == "15")
                                    {
                                        wall = w;
                                        using (Transaction t = new Transaction(doc, "Create Window"))
                                        {
                                            t.Start();
                                            if (!familySymbol.IsActive)
                                            {
                                                familySymbol.Activate();
                                                doc.Regenerate();
                                            }
                                            FamilyInstance Window = doc.Create.NewFamilyInstance(midpoint, familySymbol, wall, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                            //Parameter sill = Window.LookupParameter("Sill Height");
                                            //sill.SetValueString(offset);
                                            t.Commit();

                                        }
                                        break;

                                    }
                                }

                            }
                        }
                    }

                    TaskDialog.Show("2D-3D", "Window Created");
                }
                #endregion Window

                #region Roofs
                if (rootElement.Element("CAP").Element("Roofs") != null)
                {
                    var colroof = rootElement.Element("CAP").Element("Roofs").Elements().ToList();



                    List<XYZ> collpnts = new List<XYZ>();
                    List<string> collpntsting = new List<string>();
                    List<string[]> roofpnt = new List<string[]>();
                    CurveArray profile = new CurveArray();
                    levelvalue = 1;
                    foreach (var roofname in colroof)
                    {
                        var lstle = roofname.Element("Level_Value");
                        var lName = roofname.Element("Level_Name").Value.ToString();
                        var lNo = roofname.Element("Level_Number").Value.ToString();
                        var cPoints = roofname.Elements("cPoint");
                        var famly = roofname.Element("familyname").Value.ToString();
                        var typ = roofname.Element("type").Value.ToString();
                        var ange = roofname.Element("angle").Value.ToString();
                        var ofst = roofname.Element("offset").Value.ToString();

                        string family = famly.Replace("(", "").Replace(")", "");
                        string type = typ.Replace("(", "").Replace(")", "");
                        string angl = ange.Replace("(", "").Replace(")", "");
                        string offset = ofst.Replace("(", "").Replace(")", "");


                        Double angle = Convert.ToDouble(angl);

                        var lvlcollection = colllevel.Where(_layer => _layer.Name == "Level " + lNo).Select(i => i).ToList();
                        Level level = lvlcollection[0] as Level;

                        foreach (string cPoint in cPoints)
                        {
                            string point = cPoint.Replace("<cPoint>", "").Replace("</cPoint>", "");
                            myStrings = new[] { point };
                            roofpnt.Add(myStrings);
                        }

                        foreach (string[] star in roofpnt)
                        {
                            string roof = star[0].ToString();
                            string[] point = roof.Split(')');

                            string pnt1 = point[0].Replace("(", "");
                            string pnt2 = point[1].Replace("(", "");
                            string pnt3 = point[2].Replace("(", "");
                            string pnt4 = point[3].Replace("(", "");


                            string[] points = roof.Split(')');

                            foreach (string pnt in points)
                            {
                                string poits = pnt.Replace("(", "");
                                collpntsting.Add(poits);

                            }
                            foreach (string pnt in collpntsting)
                            {
                                string[] p1 = pnt.Split(',');
                                if (p1[0] != "")
                                {
                                    double a1 = Convert.ToDouble(p1[0]);
                                    double b1 = Convert.ToDouble(p1[1]);
                                    double c1 = Convert.ToDouble(p1[2]);
                                    XYZ ptt1 = new XYZ(a1, b1, c1);

                                    collpnts.Add(ptt1);
                                }
                            }



                            for (int i = 0; i < collpnts.Count; i++)
                            {
                                Line line = Line.CreateBound(collpnts[i],
                          collpnts[(i < collpnts.Count - 1) ? i + 1 : 0]);

                                profile.Append(line);
                            }



                            RoofType rooft = new FilteredElementCollector(doc).OfClass(typeof(RoofType)).FirstOrDefault<Element>() as RoofType;


                            Application application = doc.Application;
                            CurveArray footprint = application.Create.NewCurveArray();

                            FilteredElementCollector roofcollect = new FilteredElementCollector(doc);
                            roofcollect = new FilteredElementCollector(doc);

                            roofcollect.OfClass(typeof(RoofType));
                            //RoofType roofTypee = roofcollect.Last() as RoofType;

                            foreach (RoofType rt in roofcollect)
                            {
                                if (rt.FamilyName == family )
                                {
                                    RoofType rooftype = rt as RoofType;

                                    ModelCurveArray footPrintToModelCurveMapping = new ModelCurveArray();

                                    using (Transaction transaction = new Transaction(doc, "Place Roof"))
                                    {
                                        if (!transaction.HasStarted()) transaction.Start();

                                        try
                                        {
                                            FootPrintRoof footprintRoof = doc.Create.NewFootPrintRoof(profile, level, rooftype, out footPrintToModelCurveMapping);

                                            if (angle > 0)
                                            {
                                                foreach (ModelCurve curve in footPrintToModelCurveMapping)
                                                {
                                                    if (curve.GeometryCurve.Length > 25)
                                                    {
                                                        footprintRoof.set_DefinesSlope(curve, true);
                                                        footprintRoof.set_SlopeAngle(curve, angle);
                                                        Parameter BO = footprintRoof.LookupParameter("Base Offset From Level");
                                                        BO.SetValueString(offset);
                                                    }

                                                }
                                            }
                                            else
                                            {
                                                foreach (ModelCurve curve in footPrintToModelCurveMapping)
                                                {
                                                    footprintRoof.set_DefinesSlope(curve, true);
                                                    footprintRoof.set_SlopeAngle(curve, angle);
                                                    Parameter BO = footprintRoof.LookupParameter("Base Offset From Level");
                                                    BO.SetValueString(offset);

                                                }

                                            }

                                        }
                                        catch (Exception ex)
                                        {

                                        }
                                        transaction.Commit();
                                    }
                                }
                            }
                            collpnts.Clear();
                            collpntsting.Clear();
                            profile.Clear();



                        }
                        roofpnt.Clear();
                    }
                    TaskDialog.Show("2D-3D", "Roof Created");
                }
                #endregion Roofs

                // alternative(doc, colllevel, caplocation);
            }

            return Result.Succeeded;
        }
        public void alternative(Document doc, FilteredElementCollector colllevel, string caplocation)
        {
            Oldstartup oldstartup = new Oldstartup();

            #region i cap
            string first = string.Empty; string second = string.Empty; string _Height = string.Empty; string _Volume = string.Empty; string _Width = string.Empty; string _Function = string.Empty;
            string Cfirst = string.Empty; string Csecond = string.Empty; string C_Height = string.Empty; string C_Volume = string.Empty; string C_Width = string.Empty; string CWLevel = string.Empty;

            string Wposx = string.Empty;
            string Wposy = string.Empty;
            string Wposz = string.Empty;
            string DoorName = string.Empty;
            string FStartPoint = string.Empty;
            string FEndPoint = string.Empty;
            string FLevel = string.Empty;
            string WLevel = string.Empty;
            string wfamil = string.Empty;
            string Fwidth = string.Empty;
            string FOffset = string.Empty;

            string FSStartPoint = string.Empty;
            string FSEndPoint = string.Empty;
            string SFLevel = string.Empty;
            string SFwidth = string.Empty;
            string SFOffset = string.Empty;

            string ColumnLevel = string.Empty;

            string FillregionPoint = string.Empty;
            string CircleFilledRegion = string.Empty;
            string Circlepoints = string.Empty;
            string Midpoint = string.Empty;
            string FillLevel = string.Empty;
            string BeamStartPoint = string.Empty;
            string BeamEndPoint = string.Empty;
            string BeamLevel = string.Empty;
            string BeamWidth = string.Empty;
            string BeamHight = string.Empty;

            string doorx = string.Empty;
            string doory = string.Empty;
            string doorz = string.Empty;
            string doorlevel = string.Empty;
            string doorwidth = string.Empty;
            string doorheight = string.Empty;
            string walllength = string.Empty;


            int n = 1;

            List<string> points = new List<string>();
            IList<String[]> Floorcoll = new List<String[]>();
            IList<String[]> SFloorcoll = new List<String[]>();

            IList<String[]> wallcoll = new List<String[]>();
            IList<String[]> Cwallcoll = new List<String[]>();
            IList<String[]> Fillregio = new List<String[]>();
            IList<String[]> circleFillregio = new List<String[]>();
            IList<String[]> Beamcoll = new List<String[]>();
            IList<String[]> doorcoll = new List<String[]>();

            #region c
            //                string caploction = Path.Combine(Path.GetFullPath(System.Reflection.Assembly.GetExecutingAssembly().Location)
            //.Remove(Path.GetFullPath(System.Reflection.Assembly.GetExecutingAssembly().Location).Count()
            //- Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location).Count()), @"ECAP\2D-3D.cap");

            //using (XmlReader reader = XmlReader.Create(caplocation))
            #endregion c
            using (XmlReader reader = XmlReader.Create(caplocation))
            {
                while (reader.Read())
                {
                    if (reader.IsStartElement())
                    {
                        //return only when you have START tag
                        switch (reader.Name.ToString())
                        {
                            //FilledRegion
                            case "Points":
                                string Fillpoint = reader.ReadString();
                                FillregionPoint = Fillpoint;
                                break;

                            case "Circlepoints":
                                string Cpoints = reader.ReadString();
                                Circlepoints = Cpoints;
                                break;

                            case "Radius":
                                string Mpoint = reader.ReadString();
                                Midpoint = Mpoint;
                                break;
                            case "FillLevel":
                                string Flleel = reader.ReadString();
                                FillLevel = Flleel;
                                break;


                            //Floor
                            case "FStartPoint":
                                string Bsp = reader.ReadString();
                                FStartPoint = Bsp;
                                break;
                            case "FEndPoint":
                                string Bep = reader.ReadString();
                                FEndPoint = Bep;
                                break;
                            case "Level":
                                string lvl = reader.ReadString();
                                FLevel = lvl;
                                break;
                            case "FWidth":
                                string Fwid = reader.ReadString();
                                Fwidth = Fwid;
                                break;
                            case "Offset":
                                string ofset = reader.ReadString();
                                FOffset = ofset;
                                break;


                            case "FSStartPoint":
                                string SBsp = reader.ReadString();
                                FSStartPoint = SBsp;
                                break;
                            case "FSEndPoint":
                                string SBep = reader.ReadString();
                                FSEndPoint = SBep;
                                break;
                            case "SLevel":
                                string Slvl = reader.ReadString();
                                SFLevel = Slvl;
                                break;
                            case "SFWidth":
                                string SFwid = reader.ReadString();
                                SFwidth = SFwid;
                                break;
                            case "SOffset":
                                string Sofset = reader.ReadString();
                                SFOffset = Sofset;
                                break;


                            //Wall

                            case "StartPoint":
                                string startpoint = reader.ReadString();
                                first = startpoint;
                                break;
                            case "EndPoint":
                                string endpoint = reader.ReadString();
                                second = endpoint;
                                break;
                            case "Height":
                                string Height = reader.ReadString();
                                _Height = Height;
                                break;

                            case "Width":
                                string Width = reader.ReadString();
                                _Width = Width;
                                break;
                            case "WLevel":
                                string Wlvl = reader.ReadString();
                                WLevel = Wlvl;
                                break;
                            case "Family":
                                string wallFamily = reader.ReadString();
                                wfamil = wallFamily;
                                break;

                            //C_Wall

                            case "CStartPoint":
                                string Cstartpoint = reader.ReadString();
                                Cfirst = Cstartpoint;
                                break;
                            case "CEndPoint":
                                string Cendpoint = reader.ReadString();
                                Csecond = Cendpoint;
                                break;
                            case "CHeight":
                                string CHeight = reader.ReadString();
                                C_Height = CHeight;
                                break;

                            case "CWidth":
                                string CWidth = reader.ReadString();
                                C_Width = CWidth;
                                break;
                            case "CWLevel":
                                string CWlvl = reader.ReadString();
                                CWLevel = CWlvl;
                                break;



                            //Beam
                            case "BeamStartPoint":
                                string Bemstr = reader.ReadString();
                                BeamStartPoint = Bemstr;
                                break;
                            case "BeamEndPoint":
                                string Bemend = reader.ReadString();
                                BeamEndPoint = Bemend;
                                break;
                            case "BeamLevel":
                                string Bemlevel = reader.ReadString();
                                BeamLevel = Bemlevel;
                                break;
                            case "BeamWidth":
                                string BeamWid = reader.ReadString();
                                BeamWidth = BeamWid;
                                break;
                            case "BeamHight":
                                string BHight = reader.ReadString();
                                BeamHight = BHight;
                                break;


                            //Door
                            case "X":
                                string x = reader.ReadString();
                                doorx = x;
                                break;
                            case "Y":
                                string Y = reader.ReadString();
                                doory = Y;
                                break;
                            case "Z":
                                string z = reader.ReadString();
                                doorz = z;
                                break;
                            case "DoorLevel":
                                string DLevel = reader.ReadString();
                                doorlevel = DLevel;
                                break;
                            case "Walllength":
                                string wlen = reader.ReadString();
                                walllength = wlen;
                                break;


                            case "DoorWidth":
                                string Dwid = reader.ReadString();
                                doorwidth = Dwid;
                                break;
                            case "DoorHeight":
                                string Dhig = reader.ReadString();
                                doorheight = Dhig;
                                break;

                        }

                        if (FStartPoint.Length > 1 && FEndPoint.Length > 1 && FLevel.Length > 1 && Fwidth.Length > 0)
                        {
                            myStrings = new[] { FStartPoint, FEndPoint, FLevel, Fwidth, FOffset };
                            Floorcoll.Add(myStrings);
                            FStartPoint = string.Empty;
                            FEndPoint = string.Empty;
                            FLevel = string.Empty;
                            Fwidth = string.Empty;
                            FOffset = string.Empty;
                        }

                        if (FSStartPoint.Length > 1 && FSEndPoint.Length > 1 && SFLevel.Length > 1 && SFwidth.Length > 1)
                        {
                            myStrings = new[] { FSStartPoint, FSEndPoint, SFLevel, SFwidth, SFOffset };
                            SFloorcoll.Add(myStrings);
                            FSStartPoint = string.Empty;
                            FSEndPoint = string.Empty;
                            SFLevel = string.Empty;
                            SFwidth = string.Empty;
                            SFOffset = string.Empty;
                        }

                        if (first.Length > 1 && second.Length > 1 && _Height.Length > 1 && WLevel.Length > 1 && _Width.Length > 1)
                        {
                            myStrings = new[] { first, second, WLevel, _Height, _Width };
                            wallcoll.Add(myStrings);
                            first = string.Empty; second = string.Empty; _Height = string.Empty; _Volume = string.Empty; _Width = string.Empty; _Function = string.Empty; WLevel = string.Empty;
                        }

                        if (Cfirst.Length > 1 && Csecond.Length > 1 && C_Height.Length > 1 && CWLevel.Length > 1)
                        {
                            myStrings = new[] { Cfirst, Csecond, CWLevel, C_Height, C_Width };
                            Cwallcoll.Add(myStrings);
                            Cfirst = string.Empty; Csecond = string.Empty; C_Height = string.Empty; C_Volume = string.Empty; C_Width = string.Empty; CWLevel = string.Empty;
                        }

                        if (FillregionPoint.Length > 1)
                        {

                            myStrings = new[] { FillregionPoint };
                            Fillregio.Add(myStrings);

                            if (Fillregio.Count() == 5)
                            {
                                //createfillregion(doc, Fillregio);
                                Fillregio.Clear();
                                FillregionPoint = string.Empty;
                                ColumnLevel = string.Empty;
                            }
                        }

                        if (Circlepoints.Length > 1 && Midpoint.Length > 1 && FillLevel.Length > 1)
                        {

                            myStrings = new[] { Circlepoints, Midpoint, FillLevel };
                            circleFillregio.Add(myStrings);

                            if (circleFillregio.Count() == 1)
                            {
                                //createCirclefillregion(doc, circleFillregio);
                                circleFillregio.Clear();
                                Midpoint = string.Empty;
                                FillLevel = string.Empty;
                            }
                        }

                        if (BeamStartPoint.Length > 1 && BeamEndPoint.Length > 1 && BeamLevel.Length > 1 && BeamWidth.Length > 1 && BeamHight.Length > 1)
                        {
                            myStrings = new[] { BeamStartPoint, BeamEndPoint, BeamLevel, BeamWidth, BeamHight };
                            Beamcoll.Add(myStrings);
                            BeamStartPoint = string.Empty; BeamEndPoint = string.Empty; BeamLevel = string.Empty; BeamWidth = string.Empty; BeamHight = string.Empty;
                        }

                        if (doorx.Length > 1 && doory.Length > 1 && doorz.Length > 0 && doorlevel.Length > 1 && walllength.Length > 1 && doorwidth.Length > 1 && doorheight.Length > 1)
                        {
                            myStrings = new[] { doorx, doory, doorz, doorlevel, walllength, doorwidth, doorheight };
                            doorcoll.Add(myStrings);

                            doorx = string.Empty;
                            doory = string.Empty;
                            doorz = string.Empty;
                            doorlevel = string.Empty;
                            walllength = string.Empty;
                            doorwidth = string.Empty;
                            doorheight = string.Empty;
                        }

                    }
                }
            }
            #endregion i cap

            #region Column

            //foreach (var lvlname in colllevel)
            //{
            //    Level _l = (lvlname) as Level;

            //    createColumn(doc, _l);
            //}

            //MessageBox.Show("Column Placed", "2D-3D");
            #endregion Column

            #region Floor
            foreach (var lvlname in colllevel)
            {
                Level _l = (lvlname) as Level;

                if (Floorcoll.Count > 1)
                {
                    oldstartup.createfloor(doc, _l, Floorcoll, FLevel);
                    FEndPoint = string.Empty; FEndPoint = string.Empty; Fwidth = string.Empty; FOffset = string.Empty;
                }
                //if (SFloorcoll.Count > 1)
                //{
                //    createfloor(doc, _l, SFloorcoll, SFLevel);
                //    FSEndPoint = string.Empty; FSEndPoint = string.Empty; SFwidth = string.Empty; SFOffset = string.Empty;

                //}

            }
            System.Windows.MessageBox.Show("Floor Placed", "2D-3D");
            #endregion Floor

            #region Beam

            //foreach (var lvlname in colllevel)
            //{
            //    Level _l = (lvlname) as Level;
            //    createBeam(Beamcoll, _l, doc);
            //    BeamStartPoint = string.Empty; BeamEndPoint = string.Empty; BeamLevel = string.Empty; BeamWidth = string.Empty; BeamHight = string.Empty;
            //}
            //MessageBox.Show("Beam Placed", "2D-3D");

            #endregion Beam

            #region Wall
            foreach (var lvlname in colllevel)
            {
                Level _l = (lvlname) as Level;
                if (wallcoll.Count > 1)
                {
                    oldstartup.createwall(wallcoll, first, second, doc, _l, WLevel);
                    first = string.Empty; second = string.Empty; _Height = string.Empty; _Volume = string.Empty; _Width = string.Empty; _Function = string.Empty; WLevel = string.Empty;
                }
                //if (Cwallcoll.Count > 1)
                //{
                //    createCwall(Cwallcoll, Cfirst, Csecond, doc, _l, CWLevel);
                //    Cfirst = string.Empty; Csecond = string.Empty; C_Height = string.Empty; C_Volume = string.Empty; C_Width = string.Empty; CWLevel = string.Empty;
                //}


            }
            //wallgeomentry(doc);
            // MessageBox.Show("Wall Placed", "2D-3D");

            #endregion Wall
            #region Roof
            //foreach (var lvlname in colllevel)
            //{
            //    Level _l = (lvlname) as Level;
            //    createroof(doc, _l);
            //}


            //MessageBox.Show("Roof Placed", "2D-3D");
            #endregion Roof
            #region Door
            //foreach (var lvlname in colllevel)
            //{
            //    Level _l = (lvlname) as Level;
            //    createDoor(doorcoll, _l, doc);
            //    doorx = string.Empty; doory = string.Empty; doorz = string.Empty; doorlevel = string.Empty; walllength = string.Empty; doorwidth = string.Empty;
            //    doorheight = string.Empty;

            //}
            //MessageBox.Show("Door Placed", "2D-3D");
            #endregion Door

            TaskDialog.Show("2D-3D", "Completed");

        }
        public DataTable xlsxToDT(string path)
        {
            DataTable dt1 = new DataTable();
            try
            {
                using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    XSSFWorkbook hssfworkbook = new XSSFWorkbook(file);
                    ISheet sheet = hssfworkbook.GetSheetAt(0);
                    //  ISheet sheet = rdoElec.Checked ? hssfworkbook.GetSheet("Parcel_Elec") : hssfworkbook.GetSheet("Parcel_Mech");
                    //if (sheet == null)
                    //    sheet = hssfworkbook.GetSheetAt(0);

                    IRow headerRow = sheet.GetRow(0);
                    System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
                    int colCount = headerRow.LastCellNum;
                    int rowCount = sheet.LastRowNum;

                    for (int c = 0; c < colCount; c++)
                    {
                        try
                        {
                            dt1.Columns.Add(headerRow.GetCell(c).ToString());
                        }
                        catch (Exception ex) { dt1.Columns.Add(Guid.NewGuid().ToString()); }
                    }
                    bool skipReadingHeaderRow = rows.MoveNext();
                    while (rows.MoveNext())
                    {
                        {
                            IRow row = (XSSFRow)rows.Current;
                            DataRow dr = dt1.NewRow();
                            for (int i = 0; i < colCount; i++)
                            {

                                ICell cell = row.GetCell(i);

                                if (cell != null)
                                {
                                    if (cell.CellType == NPOI.SS.UserModel.CellType.Formula)
                                    {
                                        if (((NPOI.XSSF.UserModel.XSSFCell)cell).CachedFormulaResultType == NPOI.SS.UserModel.CellType.Numeric)
                                            dr[i] = cell.NumericCellValue.ToString();
                                        else if (((NPOI.XSSF.UserModel.XSSFCell)cell).CachedFormulaResultType == NPOI.SS.UserModel.CellType.String)
                                            dr[i] = cell.StringCellValue.ToString();
                                    }
                                    else
                                        dr[i] = cell.ToString();
                                }
                            }
                            dt1.Rows.Add(dr);
                        }
                    }
                    hssfworkbook = null;
                    sheet = null;
                }
            }
            catch (Exception ex) { System.Windows.Forms.MessageBox.Show(ex.Message.ToString()); }
            return dt1;
        }
        public static List<WallType> GetWallTypes(Autodesk.Revit.DB.Document doc)
        {
            using (Transaction transaction = new Transaction(doc, "WallType"))
            {
                if (!transaction.HasStarted()) transaction.Start();

                List<WallType> oWallTypes = new List<WallType>();
                try
                {
                    FilteredElementCollector collector
                        = new FilteredElementCollector(doc);

                    FilteredElementIterator itor = collector
                        .OfClass(typeof(HostObjAttributes))
                        .GetElementIterator();

                    // Reset the iterator
                    itor.Reset();

                    // Iterate through each family
                    while (itor.MoveNext())
                    {
                        Autodesk.Revit.DB.HostObjAttributes oSystemFamilies =
                        itor.Current as Autodesk.Revit.DB.HostObjAttributes;

                        if (oSystemFamilies == null) continue;

                        // Get the family's category
                        Category oCategory = oSystemFamilies.Category;

                        // Process if the category is found
                        if (oCategory != null)
                        {
                            if (oCategory.Name == "Walls")
                            {
                                WallType oWallType = oSystemFamilies as WallType;
                                if (oWallType != null) oWallTypes.Add(oWallType);
                            }
                        }
                    } //while itor.NextMove()

                    transaction.Commit();
                    return oWallTypes;
                }
                catch (Exception)
                {
                    //MessageBox.Show( ex.Message );
                    return oWallTypes = new List<WallType>();
                }
            }
        }
        public static List<FloorType> collfloorTypes(Autodesk.Revit.DB.Document doc)
        {
            using (Transaction transaction = new Transaction(doc, "floortype"))
            {
                if (!transaction.HasStarted()) transaction.Start();

                List<FloorType> ofloorTypes = new List<FloorType>();
                try
                {
                    FilteredElementCollector collector
                        = new FilteredElementCollector(doc);

                    FilteredElementIterator itor = collector
                        .OfClass(typeof(HostObjAttributes))
                        .GetElementIterator();

                    // Reset the iterator
                    itor.Reset();

                    // Iterate through each family
                    while (itor.MoveNext())
                    {
                        Autodesk.Revit.DB.HostObjAttributes oSystemFamilies =
                        itor.Current as Autodesk.Revit.DB.HostObjAttributes;

                        if (oSystemFamilies == null) continue;

                        // Get the family's category
                        Category oCategory = oSystemFamilies.Category;

                        // Process if the category is found
                        if (oCategory != null)
                        {
                            if (oCategory.Name == "Floors")
                            {
                                FloorType ofloorType = oSystemFamilies as FloorType;
                                if (ofloorType != null) ofloorTypes.Add(ofloorType);
                            }
                        }
                    } //while itor.NextMove()

                    transaction.Commit();
                    return ofloorTypes;
                }
                catch (Exception)
                {
                    //MessageBox.Show( ex.Message );
                    return ofloorTypes = new List<FloorType>();
                }
            }
        }
        private static void EncFile(string filename)
        {
            try
            {
                string text = File.ReadAllText(filename);
                string EncryptionKey = "CAPINFO";
                byte[] clearBytes = Encoding.Unicode.GetBytes(text);
                using (Aes encryptor = Aes.Create())
                {
                    Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] {
            0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76});
                    encryptor.Key = pdb.GetBytes(32);
                    encryptor.IV = pdb.GetBytes(16);
                    using (MemoryStream ms = new MemoryStream())
                    {
                        using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write))
                        {
                            cs.Write(clearBytes, 0, clearBytes.Length);
                            cs.Close();
                        }
                        text = Convert.ToBase64String(ms.ToArray());
                    }
                }
                using (StreamWriter writetext = new StreamWriter(filename))
                {
                    writetext.WriteLine(text);
                }
            }
            catch (Exception ex) { System.Windows.MessageBox.Show("ExportCap", "" + ex.Message + ex.StackTrace); }

        }
        private static void DecFile(string filename)
        {
            try
            {

                string text = File.ReadAllText(filename);
                string EncryptionKey = "CAPINFO";
                text = text.Replace(" ", "+");
                byte[] cipherBytes = Convert.FromBase64String(text);
                using (Aes encryptor = Aes.Create())
                {
                    Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] {
            0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76});
                    encryptor.Key = pdb.GetBytes(32);
                    encryptor.IV = pdb.GetBytes(16);
                    using (MemoryStream ms = new MemoryStream())
                    {
                        using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write))
                        {
                            cs.Write(cipherBytes, 0, cipherBytes.Length);
                            cs.Close();
                        }
                        text = Encoding.Unicode.GetString(ms.ToArray());


                        using (StreamWriter writetext = new StreamWriter(filename))
                        {
                            writetext.WriteLine(text);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("ExportCap", "" + ex.Message + ex.StackTrace);
            }
        }
        public static void deletelevel(Document document)
        {
            int deleted = 0;
            var colllevel = new FilteredElementCollector(document).OfClass(typeof(Level)).OfCategory(BuiltInCategory.OST_Levels);

            List<ElementId> elementsToBeDeleted = new List<ElementId>();
            using (Transaction transaction = new Transaction(document, "Place"))
            {
                if (!transaction.HasStarted()) transaction.Start();

                foreach (Element element in colllevel)
                {
                    elementsToBeDeleted.Add(element.Id);
                    deleted++;
                }
                document.Delete(elementsToBeDeleted);
                transaction.Commit();
            }
        }
        public class MyPreProcessor : IFailuresPreprocessor
        {
            FailureProcessingResult IFailuresPreprocessor.PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                String transactionName = failuresAccessor.GetTransactionName();

                IList<FailureMessageAccessor> fmas = failuresAccessor.GetFailureMessages();

                if (fmas.Count == 0)
                    return FailureProcessingResult.Continue;

                if (transactionName.Equals("EXEMPLE"))
                {
                    foreach (FailureMessageAccessor fma in fmas)
                    {
                        if (fma.GetSeverity() == FailureSeverity.Error)
                        {
                            failuresAccessor.DeleteAllWarnings();
                            return FailureProcessingResult.ProceedWithRollBack;
                        }
                        else
                        {
                            failuresAccessor.DeleteWarning(fma);
                        }

                    }
                }
                else
                {
                    foreach (FailureMessageAccessor fma in fmas)
                    {
                        failuresAccessor.DeleteAllWarnings();
                    }
                }
                return FailureProcessingResult.Continue;
            }
        }
    }
}
