using _23DC.Properties;
using Autodesk.AutoCAD.Customization;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Diagnostics;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows.BlockStream;
using GrapeCity.Documents.Pdf;
using GrapeCity.Documents.Pdf.Annotations;
using GrapeCity.Documents.Pdf.Recognition.Structure;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Xps.Packaging;
using System.Xml;
using System.Xml.Linq;
using WinCopies.Util;
using XDrawerLib;
using XDrawerLib.Drawers;
using XDrawerLib.Helpers;
using DoubleCollection = Autodesk.AutoCAD.Geometry.DoubleCollection;
using Ellipse = Autodesk.AutoCAD.DatabaseServices.Ellipse;
using Exception = System.Exception;
using Line = Autodesk.AutoCAD.DatabaseServices.Line;
using Math = System.Math;
using MessageBox = System.Windows.MessageBox;
using OpenMode = Autodesk.AutoCAD.DatabaseServices.OpenMode;
using Path = System.IO.Path;
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;

namespace _23DC
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static string layerNAME;
        public static CAP _BIMEngine = new CAP();
        public static SortedList<string, List<Curve>> levelHatchs = new SortedList<string, List<Curve>>();
        public static SortedList<string, List<Curve>> wallPolyline = new SortedList<string, List<Curve>>();
        public static List<double> lstLevels = new List<double>();
        public static List<string> lstLevelAreas = new List<string>();
        public static List<double> lstWallLength = new List<double>();
        public static string currentFilename = string.Empty;
        public static bool isSection = false;
        public static bool isElevation = false;
        public static bool isFloor = false;
        public static int countHatchs = 0;
        public static double slidervalue = 10;
        private static List<string> lstFiles = new List<string>();
        private int filecount;
        public static string filename = string.Empty;
        public static string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        public static string _outputcapfilename = path + System.IO.Path.DirectorySeparatorChar + "Temp" + System.IO.Path.DirectorySeparatorChar + "dkn.cap";
        public static string dlllocation = path + System.IO.Path.DirectorySeparatorChar + "ECAP.dll";

        public MainWindow()
        {
            InitializeComponent();
            txtBrowse.Text = Properties.Settings.Default.catchEcapPath;
            txtBrowse_Output.Text = Properties.Settings.Default.catchIcapPath;
            //var uriSource = new Uri(@"pack://application:,,,/23DC;component/Resources/magnifying_glass.png", UriKind.Relative);
            //ico_btnScan.Source = new BitmapImage(uriSource);
            lstFiles = Directory.GetFiles(txtBrowse.Text, "*.dwg*").ToList();

        }

        // Functions

        #region Functions

        private bool IsFileWritable(string filename)
        {
            if (!File.GetAttributes(filename).HasFlag(FileAttributes.ReadOnly))
            {
                try
                {
                    using (new FileInfo(filename).Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    {
                        return true;
                    }
                }
                catch (IOException) { }
            }
            return false;
        }
        private static void xportAcDbAlignedDimension(List<ObjectId> acDbAlignedDimension, XmlWriter writer, OpenCloseTransaction tr)
        {
            try
            {
                foreach (var EntName in acDbAlignedDimension)
                {
                    string content = string.Empty;
                    Entity enttxt = (Entity)tr.GetObject(EntName, OpenMode.ForWrite);
                    if (enttxt is AlignedDimension)
                    {
                        AlignedDimension leader = ((AlignedDimension)enttxt);

                        var elename = EntName.ObjectClass.Name;
                        var leaderWeight = leader.Database.Dimlwd;

                        var DimLinePoint = ((Autodesk.AutoCAD.DatabaseServices.AlignedDimension)enttxt).DimLinePoint;
                        var DimStyle = leader.DimensionStyleName;
                        var XLine1Point = leader.XLine1Point;
                        var XLine2Point = leader.XLine2Point;
                        var meterial = leader.Material;
                        layerNAME = leader.Layer;

                        writer.WriteStartElement("AlignedDimension");
                        writer.WriteAttributeString("ElementID", "" + EntName.ToString());
                        writer.WriteAttributeString("ElementName", "" + elename.ToString());
                        writer.WriteAttributeString("LayerName", "" + layerNAME.ToString());
                        writer.WriteAttributeString("LeaderWeight", "" + leaderWeight.ToString());
                        writer.WriteAttributeString("DimLinePoint", "" + DimLinePoint.ToString());
                        writer.WriteAttributeString("DimStyle", "" + DimStyle.ToString());
                        writer.WriteAttributeString("Material", "" + meterial.ToString());
                        writer.WriteElementString("XLine1Point", "" + XLine1Point.ToString());
                        writer.WriteElementString("XLine2Point", "" + XLine2Point.ToString());

                        writer.WriteEndElement();
                    }

                }
            }
            catch (Exception ex) { MessageBox.Show("", "" + ex.Message); }
        }
        //export Rotated Dimention
        private static void xportAcDbRotatedDimension(List<ObjectId> acDbRotatedDimension, XmlWriter writer, OpenCloseTransaction tr)
        {
            try
            {
                foreach (var EntName in acDbRotatedDimension)
                {
                    string content = string.Empty;
                    Entity enttxt = (Entity)tr.GetObject(EntName, OpenMode.ForWrite);
                    if (enttxt is RotatedDimension)
                    {
                        RotatedDimension leader = ((RotatedDimension)enttxt);

                        var elename = EntName.ObjectClass.Name;
                        var leaderWeight = leader.Database.Dimlwd;
                        var DimStyle = leader.DimensionStyleName;
                        var TextPosition = leader.TextPosition;
                        var Text = leader.DimensionText;
                        var XLine1Point = leader.XLine1Point;
                        var XLine2Point = leader.XLine2Point;
                        var meterial = leader.Material;
                        layerNAME = leader.Layer;

                        if (Text.Length >= 1)
                        {
                            writer.WriteStartElement("RotatedDimension");
                            writer.WriteAttributeString("ElementID", "" + EntName.ToString());
                            writer.WriteAttributeString("ElementName", "" + elename.ToString());
                            writer.WriteAttributeString("LayerName", "" + layerNAME.ToString());
                            writer.WriteAttributeString("LeaderWeight", "" + leaderWeight.ToString());
                            writer.WriteAttributeString("DimStyle", "" + DimStyle.ToString());
                            writer.WriteAttributeString("TextPosition", "" + TextPosition.ToString());
                            writer.WriteAttributeString("Material", "" + meterial.ToString());
                            writer.WriteElementString("Content", "" + Text.ToString());
                            writer.WriteElementString("XLine1Point", "" + XLine1Point.ToString());
                            writer.WriteElementString("XLine2Point", "" + XLine2Point.ToString());

                            writer.WriteEndElement();
                        }
                    }

                }
            }
            catch (Exception ex) { MessageBox.Show("", "" + ex.Message); }
        }
        //export Leader 
        private static void xportAcDbLeader(List<ObjectId> acDbLeader, XmlWriter writer, OpenCloseTransaction tr)
        {
            try
            {
                foreach (var EntName in acDbLeader)
                {
                    string content = string.Empty;
                    Entity enttxt = (Entity)tr.GetObject(EntName, OpenMode.ForWrite);
                    if (enttxt is Leader)
                    {
                        Leader leader = ((Leader)enttxt);

                        var elename = EntName.ObjectClass.Name;
                        var leaderWeight = leader.Database.Dimlwd;
                        var CurrentVertex = leader.FirstVertex;
                        var DimStyle = leader.DimensionStyleName;
                        var Type = leader.AnnoType;
                        var Annotative = leader.Annotative;
                        layerNAME = leader.Layer;


                        var TextOffset = leader.AnnotationOffset;
                        var DimScale = leader.LinetypeScale;
                        var startpoint = leader.StartPoint;
                        var Endpoint = leader.EndPoint;
                        var meterial = leader.Material;

                        writer.WriteStartElement("Leader");
                        writer.WriteAttributeString("ElementID", "" + EntName.ToString());
                        writer.WriteAttributeString("ElementName", "" + elename.ToString());
                        writer.WriteAttributeString("LayerName", "" + layerNAME.ToString());
                        writer.WriteAttributeString("LeaderWeight", "" + leaderWeight.ToString());
                        writer.WriteAttributeString("CurrentVertex", "" + CurrentVertex.ToString());
                        writer.WriteAttributeString("DimStyle", "" + DimStyle.ToString());
                        writer.WriteAttributeString("Type", "" + Type.ToString());
                        writer.WriteAttributeString("Annotative", "" + Annotative.ToString());
                        writer.WriteAttributeString("TextOffset", "" + TextOffset.ToString());
                        writer.WriteAttributeString("DimScale", "" + DimScale.ToString());
                        writer.WriteAttributeString("Material", "" + meterial.ToString());
                        writer.WriteElementString("startpoint", "" + startpoint.ToString());
                        writer.WriteElementString("Endpoint", "" + Endpoint.ToString());

                        writer.WriteEndElement();
                    }

                }
            }
            catch (Exception ex) { MessageBox.Show("", "" + ex.Message); }
        }
        //xport Ellipse
        private static void xportAcDbEllipse(List<ObjectId> acDbEllipse, XmlWriter writer, OpenCloseTransaction tr)
        {
            try
            {
                foreach (var EntName in acDbEllipse)
                {
                    string content = string.Empty;
                    Entity enttxt = (Entity)tr.GetObject(EntName, OpenMode.ForWrite);
                    if (enttxt is Ellipse)
                    {
                        Ellipse ellipse = ((Ellipse)enttxt);
                        var StartPoint = ellipse.StartPoint;
                        var EndPoint = ellipse.EndPoint;
                        var Center = ellipse.Center;
                        var MajorRadius = ellipse.MajorRadius;
                        var MinorRadius = ellipse.MinorRadius;
                        var RadiusRatio = ellipse.RadiusRatio;
                        var StartAngle = ellipse.StartAngle;
                        var EndAngle = ellipse.EndAngle;
                        var MajorAxis = ellipse.MajorAxis;
                        var MinorAxis = ellipse.MinorAxis;
                        var Area = ellipse.Area;
                        layerNAME = ellipse.Layer;
                        var elename = EntName.ObjectClass.Name;
                        var ellipseweight = ellipse.LineWeight;
                        var meterial = ellipse.Material;

                        writer.WriteStartElement("Ellipse");
                        writer.WriteAttributeString("ElementID", "" + EntName.ToString());
                        writer.WriteAttributeString("ElementName", "" + elename.ToString());
                        writer.WriteAttributeString("LayerName", "" + layerNAME.ToString());
                        writer.WriteAttributeString("Center", "" + Center.ToString());
                        writer.WriteAttributeString("MajorRadius", "" + MajorRadius.ToString());
                        writer.WriteAttributeString("MinorRadius", "" + MinorRadius.ToString());
                        writer.WriteAttributeString("RadiusRatio", "" + RadiusRatio.ToString());
                        writer.WriteAttributeString("StartAngle", "" + StartAngle.ToString());
                        writer.WriteAttributeString("EndAngle", "" + EndAngle.ToString());
                        writer.WriteAttributeString("MajorAxis", "" + MajorAxis.ToString());
                        writer.WriteAttributeString("MinorAxis", "" + MinorAxis.ToString());
                        writer.WriteAttributeString("Area", "" + Area.ToString());
                        writer.WriteAttributeString("ellipseweight", "" + ellipseweight.ToString());
                        writer.WriteAttributeString("Material", "" + meterial.ToString());
                        writer.WriteElementString("StartPoint", "" + StartPoint.ToString());
                        writer.WriteElementString("EndPoint", "" + EndPoint.ToString());
                        writer.WriteEndElement();
                    }

                }
            }
            catch (Exception ex) { MessageBox.Show("", "" + ex.Message); }
        }
        //export hatch
        private static void xportAcDbHatch(List<ObjectId> acDbHatch, XmlWriter writer, OpenCloseTransaction tr)
        {
            try
            {
                foreach (var EntName in acDbHatch)
                {
                    string content = string.Empty;
                    Entity enttxt = (Entity)tr.GetObject(EntName, OpenMode.ForWrite);
                    if (enttxt is Hatch)
                    {
                        Hatch hatch = ((Hatch)enttxt);
                        var elename = EntName.ObjectClass.Name;
                        var type = hatch.PatternType;
                        var Maximun = hatch.GeometricExtents.MaxPoint;
                        var Minimum = hatch.GeometricExtents.MinPoint;
                        var spacing = hatch.PatternSpace;
                        var Hatchweight = hatch.LineWeight;
                        var meterial = hatch.Material;
                        var area = hatch.Area;
                        string areastr = Math.Round(area, 0).ToString();
                        List<Curve> curves = GetHatchBoundary(hatch);
                        layerNAME = hatch.Layer;
                        countHatchs++;
                        if (isFloor && areastr.Length > 7 && !lstLevelAreas.Contains(areastr + currentFilename.ToLower().Replace(" ", "")))
                        {
                            lstLevelAreas.Add(areastr + currentFilename.ToLower().Replace(" ", ""));
                            levelHatchs.Add("Hatch" + countHatchs.ToString() + ";" + currentFilename.ToLower().Replace(" ", ""), curves);
                        }
                        //pattendtype
                        writer.WriteStartElement("Hatch");
                        writer.WriteAttributeString("ElementID", "" + EntName.ToString());
                        writer.WriteAttributeString("ElementName", "" + elename.ToString());
                        writer.WriteAttributeString("LayerName", "" + layerNAME.ToString());
                        writer.WriteAttributeString("Type", "" + type.ToString());
                        writer.WriteAttributeString("Space", "" + spacing.ToString());
                        writer.WriteAttributeString("HatchWeight", "" + Hatchweight.ToString());
                        writer.WriteAttributeString("Material", "" + meterial.ToString());
                        writer.WriteAttributeString("Area", "" + area.ToString());
                        writer.WriteElementString("Maximun", "" + Maximun.ToString());
                        writer.WriteElementString("Minimum", "" + Minimum.ToString());
                        writer.WriteStartElement("Boundary");
                        int xcount = 0;
                        foreach (var curve in curves)
                        {
                            writer.WriteElementString("cPoint", "" + curve.StartPoint.ToString());
                            xcount++;
                            if (xcount == curves.Count)
                                writer.WriteElementString("cPoint", "" + curve.EndPoint.ToString());
                        }
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                    }

                }
            }
            catch (Exception ex) { MessageBox.Show("", "" + ex.Message); }
        }
        // export Mleader
        private static void xportAcDbMLeader(List<ObjectId> acDbMLeader, XmlWriter writer, OpenCloseTransaction tr)
        {
            try
            {
                foreach (var EntName in acDbMLeader)
                {
                    string content = string.Empty;
                    Entity enttxt = (Entity)tr.GetObject(EntName, OpenMode.ForWrite);
                    if (enttxt is MLeader)
                    {
                        MLeader leader = ((MLeader)enttxt);
                        var elename = EntName.ObjectClass.Name;
                        var leaderWeight = leader.Database.Dimlwd;

                        //var CurrentVertex = leader.FirstVertex;
                        //var DimStyle = leader.DimensionStyle;
                        //var Type = leader.AnnoType;
                        //var Annotative = leader.Annotative;

                        //var DimLineColor = leader.Color;
                        //var TextOffset = leader.AnnotationOffset;
                        //var DimScale = leader.LinetypeScale;
                        //var startpoint = leader.StartPoint;
                        //var Endpoint = leader.EndPoint;
                        var meterial = leader.Material;
                        layerNAME = leader.Layer;

                        writer.WriteStartElement("MLeader");
                        writer.WriteAttributeString("ElementID", "" + EntName.ToString());
                        writer.WriteAttributeString("ElementName", "" + elename.ToString());
                        writer.WriteAttributeString("LayerName", "" + layerNAME.ToString());
                        writer.WriteAttributeString("LeaderWeight", "" + leaderWeight.ToString());
                        //writer.WriteElementString("CurrentVertex", "" + CurrentVertex.ToString());
                        //writer.WriteElementString("DimStyle", "" + DimStyle.ToString());
                        //writer.WriteElementString("Type", "" + Type.ToString());
                        //writer.WriteElementString("Annotative", "" + Annotative.ToString());

                        //writer.WriteElementString("DimLineColor", "" + DimLineColor.ToString());
                        //writer.WriteElementString("TextOffset", "" + TextOffset.ToString());
                        //writer.WriteElementString("DimScale", "" + DimScale.ToString());
                        //writer.WriteElementString("startpoint", "" + startpoint.ToString());
                        //writer.WriteElementString("Endpoint", "" + Endpoint.ToString());
                        writer.WriteAttributeString("Material", "" + meterial.ToString());
                        writer.WriteEndElement();
                    }

                }
            }
            catch (Exception ex) { MessageBox.Show("", "" + ex.Message); }
        }
        //export arc
        private static void xportAcDbArc(List<ObjectId> acDbArc, XmlWriter writer, OpenCloseTransaction tr)
        {
            try
            {
                foreach (var EntName in acDbArc)
                {
                    string content = string.Empty;
                    Entity enttxt = (Entity)tr.GetObject(EntName, OpenMode.ForWrite);
                    if (enttxt is Arc)
                    {
                        Arc _Arc = ((Arc)enttxt);
                        var Radius = _Arc.Radius;
                        var StartAngle = _Arc.StartAngle;
                        var EndAngle = _Arc.EndAngle;
                        var totalAngle = _Arc.TotalAngle;
                        var _Arclength = _Arc.Length;
                        var Area = _Arc.Area;
                        layerNAME = _Arc.Layer;
                        var startpoint = _Arc.StartPoint;
                        var endpoint = _Arc.EndPoint;
                        var center = _Arc.Center;
                        var elename = EntName.ObjectClass.Name;
                        var _Arcweight = _Arc.LineWeight;

                        var meterial = _Arc.Material;

                        writer.WriteStartElement("Arc");
                        writer.WriteAttributeString("ElementID", "" + EntName.ToString());
                        writer.WriteAttributeString("ElementName", "" + elename.ToString());
                        writer.WriteAttributeString("LayerName", "" + layerNAME.ToString());
                        writer.WriteAttributeString("Radius", "" + Radius.ToString());
                        writer.WriteAttributeString("StartAngle", "" + StartAngle.ToString());
                        writer.WriteAttributeString("EndAngle", "" + EndAngle.ToString());
                        writer.WriteAttributeString("TotalAngle", "" + totalAngle.ToString());
                        writer.WriteAttributeString("ArcLenght", "" + _Arclength.ToString());
                        writer.WriteAttributeString("center", "" + center.ToString());
                        writer.WriteAttributeString("Area", "" + Area.ToString());
                        writer.WriteAttributeString("LineWeight", "" + _Arcweight.ToString());
                        writer.WriteAttributeString("Material", "" + meterial.ToString());
                        writer.WriteElementString("StartPoint", "" + startpoint.ToString());
                        writer.WriteElementString("EndPoint", "" + endpoint.ToString());
                        writer.WriteEndElement();
                    }

                }
            }
            catch (Exception ex) { MessageBox.Show("", "" + ex.Message); }
        }
        //export circle
        private static void xportAcDbCircle(List<ObjectId> acDbCircle, XmlWriter writer, OpenCloseTransaction tr)
        {
            try
            {
                foreach (var EntName in acDbCircle)
                {
                    string content = string.Empty;
                    Entity enttxt = (Entity)tr.GetObject(EntName, OpenMode.ForWrite);
                    if (enttxt is Circle)
                    {
                        Circle circle = ((Circle)enttxt);
                        var Radius = circle.Radius;
                        var diameter = circle.Diameter;
                        var circumferance = circle.Circumference;
                        var Area = circle.Area;
                        var layerNAME = circle.Layer;
                        var startpoint = circle.StartPoint;
                        var endpoint = circle.EndPoint;
                        var elename = EntName.ObjectClass.Name;
                        var circleweight = circle.LineWeight;

                        var meterial = circle.Material;

                        writer.WriteStartElement("Circle");
                        writer.WriteAttributeString("ElementID", "" + EntName.ToString());
                        writer.WriteAttributeString("ElementName", "" + elename.ToString());
                        writer.WriteAttributeString("LayerName", "" + layerNAME.ToString());
                        writer.WriteAttributeString("Radius", "" + Radius.ToString());
                        writer.WriteAttributeString("Diameter", "" + diameter.ToString());
                        writer.WriteAttributeString("Circumference", "" + circumferance.ToString());
                        writer.WriteAttributeString("Area", "" + Area.ToString());
                        writer.WriteAttributeString("LineWeight", "" + circleweight.ToString());
                        writer.WriteAttributeString("Material", "" + meterial.ToString());
                        writer.WriteElementString("StartPoint", "" + startpoint.ToString());
                        writer.WriteElementString("EndPoint", "" + endpoint.ToString());
                        writer.WriteEndElement();
                    }

                }
            }
            catch (Exception ex) { MessageBox.Show("", "" + ex.Message); }
        }
        //Export Text 
        private static void xportdbText(List<ObjectId> acDbText, XmlWriter writer, OpenCloseTransaction tr)
        {
            try
            {
                List<string> contents = new List<string>();
                foreach (var _text in acDbText)
                {
                    string content = string.Empty;
                    Entity enttxt = (Entity)tr.GetObject(_text, OpenMode.ForWrite);
                    if (enttxt is DBText)
                    {

                        DBText Text = ((DBText)enttxt);

                        content = ((DBText)enttxt).TextString;
                        layerNAME = Text.Layer;
                        var textheight = Text.Height;
                        var Location = Text.Position;
                        var elename = _text.ObjectClass.Name;
                        var meterial = Text.Material;

                        if (content.Length >= 1)
                        {
                            writer.WriteStartElement("Text");
                            writer.WriteAttributeString("ElementID", "" + _text.ToString());
                            writer.WriteAttributeString("ElementName", "" + elename.ToString());
                            writer.WriteAttributeString("LayerName", "" + layerNAME.ToString());
                            writer.WriteAttributeString("TextHeight", "" + textheight.ToString());
                            writer.WriteAttributeString("Material", "" + meterial.ToString());
                            writer.WriteAttributeString("location", "" + Location.ToString());
                            writer.WriteElementString("Content", "" + content.ToString());
                            writer.WriteEndElement();
                        }

                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("", "" + ex.Message); }
        }
        //export Mtext
        private static void xportMtext(List<ObjectId> acDbMText, XmlWriter writer, OpenCloseTransaction tr)
        {
            try
            {

                List<string> contents = new List<string>();
                foreach (var mtext in acDbMText)
                {
                    string content = string.Empty;
                    Entity enttxt = (Entity)tr.GetObject(mtext, OpenMode.ForWrite);
                    if (enttxt is MText)
                    {
                        MText Text = ((MText)enttxt);

                        var _txt = Text.Text;
                        layerNAME = Text.Layer;
                        var Location = Text.Location;
                        var elename = mtext.ObjectClass.Name;

                        var meterial = Text.Material;

                        if (_txt.Length > 1)
                        {
                            content = ((MText)enttxt).Text;
                            writer.WriteStartElement("MText");
                            writer.WriteAttributeString("ElementID", "" + mtext.ToString());
                            writer.WriteAttributeString("ElementName", "" + elename.ToString());
                            writer.WriteAttributeString("LayerName", "" + elename.ToString());
                            writer.WriteAttributeString("Material", "" + meterial.ToString());
                            writer.WriteAttributeString("location", "" + Location.ToString());
                            writer.WriteElementString("Content", "" + content.ToString());

                            writer.WriteEndElement();
                        }



                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("", "" + ex.Message); }
        }
        //export Mline
        private static void xportAcDbMline(List<ObjectId> acDbMline, XmlWriter writer, OpenCloseTransaction tr)
        {
            try
            {
                foreach (var EntName in acDbMline)
                {
                    string content = string.Empty;
                    Entity enttxt = (Entity)tr.GetObject(EntName, OpenMode.ForWrite);
                    if (enttxt is Mline)
                    {
                        Mline line = ((Mline)enttxt);

                        //var lengh = line.le;
                        var layerNAME = line.Layer;
                        //var startpoint = line.StartPoint;
                        //var endpoint = line.EndPoint;
                        var elename = EntName.ObjectClass.Name;
                        var lineweight = line.LineWeight;
                        var lineType = line;
                        var meterial = line.Material;

                        writer.WriteStartElement("Mline");
                        writer.WriteAttributeString("ElementID", "" + EntName.ToString());
                        writer.WriteAttributeString("ElementName", "" + elename.ToString());
                        writer.WriteAttributeString("LayerName", "" + layerNAME.ToString());
                        //writer.WriteElementString("Length", "" + lengh.ToString());
                        writer.WriteAttributeString("LineWeight", "" + lineweight.ToString());
                        writer.WriteAttributeString("LineType", "" + lineType.ToString());
                        writer.WriteAttributeString("Material", "" + meterial.ToString());
                        //writer.WriteElementString("StartPoint", "" + startpoint.ToString());
                        //writer.WriteElementString("EndPoint", "" + endpoint.ToString());
                        writer.WriteEndElement();
                    }

                }
            }
            catch (Exception ex) { MessageBox.Show("", "" + ex.Message); }
        }
        private static void xportSpline(List<ObjectId> acDbSpline, XmlWriter writer, OpenCloseTransaction tr)
        {
            try
            {
                foreach (var EntName in acDbSpline)
                {
                    string content = string.Empty;
                    Entity enttxt = (Entity)tr.GetObject(EntName, OpenMode.ForWrite);
                    if (enttxt is Spline)
                    {
                        Spline line = ((Spline)enttxt);


                        layerNAME = line.Layer;
                        var startpoint = line.StartPoint;
                        var endpoint = line.EndPoint;
                        var elename = EntName.ObjectClass.Name;
                        var lineweight = line.LineWeight;
                        var lineType = line.Linetype;
                        var meterial = line.Material;

                        writer.WriteStartElement("Spline");
                        writer.WriteAttributeString("ElementID", "" + EntName.ToString());
                        writer.WriteAttributeString("ElementName", "" + elename.ToString());
                        writer.WriteAttributeString("LayerName", "" + layerNAME.ToString());
                        writer.WriteAttributeString("LineWeight", "" + lineweight.ToString());
                        writer.WriteAttributeString("LineType", "" + lineType.ToString());
                        writer.WriteAttributeString("Material", "" + meterial.ToString());
                        writer.WriteElementString("StartPoint", "" + startpoint.ToString());
                        writer.WriteElementString("EndPoint", "" + endpoint.ToString());
                        writer.WriteEndElement();
                    }

                }
            }
            catch (Exception ex) { MessageBox.Show("", "" + ex.Message); }
        }
        //export polyline
        private static void xportPolyline(List<ObjectId> acDbPolyline, XmlWriter writer, OpenCloseTransaction tr)
        {
            try
            {
                double bigPolyline = 0;
                Polyline pline = null;
                List<Curve> lstcurve = new List<Curve>();
                foreach (var EntName in acDbPolyline)
                {
                    string content = string.Empty;
                    Entity enttxt = (Entity)tr.GetObject(EntName, OpenMode.ForWrite);
                    if (enttxt is Polyline)
                    {
                        Polyline line = ((Polyline)enttxt);
                        var lengh = line.Length;
                        layerNAME = line.Layer;
                        var startpoint = line.StartPoint;
                        var endpoint = line.EndPoint;
                        var elename = EntName.ObjectClass.Name;
                        var lineweight = line.LineWeight;
                        var lineType = line.Linetype;
                        var meterial = line.Material;

                        if (bigPolyline < lengh)
                        {
                            bigPolyline = lengh;
                            pline = line;
                            lstcurve = new List<Curve>();
                        }

                        lstWallLength.Add(lengh);
                        writer.WriteStartElement("Polyline");
                        writer.WriteAttributeString("ElementID", "" + EntName.ToString());
                        writer.WriteAttributeString("ElementName", "" + elename.ToString());
                        writer.WriteAttributeString("LayerName", "" + layerNAME.ToString());
                        writer.WriteAttributeString("Length", "" + lengh.ToString());
                        writer.WriteAttributeString("LineWeight", "" + lineweight.ToString());
                        writer.WriteAttributeString("LineType", "" + lineType.ToString());
                        writer.WriteAttributeString("Material", "" + meterial.ToString());


                        writer.WriteElementString("StartPoint", "" + startpoint.ToString());
                        writer.WriteElementString("EndPoint", "" + endpoint.ToString());

                        writer.WriteStartElement("Boundary");
                        int xcount = 0;
                        using (var curves = new DBObjectCollection())
                        {
                            pline.Explode(curves);
                            foreach (var curve in curves)
                            {
                                if (curve is Curve)
                                {
                                    if (bigPolyline == lengh)
                                        lstcurve.Add(((Curve)curve));
                                    writer.WriteElementString("cPoint", "" + ((Curve)curve).StartPoint.ToString());
                                    xcount++;
                                    if (xcount == curves.Count)
                                        writer.WriteElementString("cPoint", "" + ((Curve)curve).EndPoint.ToString());
                                }
                            }
                        }
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                    }

                }
                if (isFloor)
                {
                    wallPolyline.Add("Polyline" + countHatchs.ToString() + ";" + currentFilename.ToLower().Replace(" ", ""), lstcurve);
                }
            }
            catch (Exception ex) { MessageBox.Show("", "" + ex.Message); }
        }
        // exportline details
        private static void Xportline(List<ObjectId> acDbLine, System.Xml.XmlWriter writer, OpenCloseTransaction tr)
        {
            try
            {
                foreach (var EntName in acDbLine)
                {
                    string content = string.Empty;
                    Entity enttxt = (Entity)tr.GetObject(EntName, OpenMode.ForWrite);
                    if (enttxt is Line)
                    {
                        Line line = ((Line)enttxt);
                        var lengh = line.Length;
                        var layerNAME = line.Layer;
                        var startpoint = line.StartPoint;
                        var endpoint = line.EndPoint;
                        var elename = EntName.ObjectClass.Name;
                        var lineweight = line.LineWeight;
                        var lineType = line.Angle;
                        var meterial = line.Material;

                        writer.WriteStartElement("Line");
                        writer.WriteAttributeString("ElementID", "" + EntName.ToString());
                        writer.WriteAttributeString("ElementName", "" + elename.ToString());
                        writer.WriteAttributeString("LayerName", "" + layerNAME.ToString());
                        writer.WriteAttributeString("Length", "" + lengh.ToString());
                        writer.WriteAttributeString("LineWeight", "" + lineweight.ToString());
                        writer.WriteAttributeString("LineType", "" + lineType.ToString());
                        writer.WriteAttributeString("Material", "" + meterial.ToString());
                        writer.WriteElementString("StartPoint", "" + startpoint.ToString());
                        writer.WriteElementString("EndPoint", "" + endpoint.ToString());
                        writer.WriteEndElement();
                    }

                }
            }
            catch (Exception ex) { MessageBox.Show("", "" + ex.Message); }
        }
        //UnlockLayer
        private static void UlockLayer(OpenCloseTransaction tr, Database db)
        {
            try
            {
                LayerTable acLyrTbl = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                foreach (var ltrId in acLyrTbl)
                {
                    if (ltrId != db.Clayer && (ltrId != db.LayerZero))
                    {
                        var ltr = (LayerTableRecord)tr.GetObject(ltrId, OpenMode.ForWrite);
                        ltr.IsLocked = false;
                        ltr.IsOff = ltr.IsOff;
                    }
                }
            }
            catch (Exception ex)
            { MessageBox.Show("EXportCAP Failed", "" + ex.Message + ex.StackTrace); }
        }
        private static bool findText(List<ObjectId> acDbText, OpenCloseTransaction tr, string _findtxt)
        {
            try
            {
                List<string> lstresultDbText = new List<string>();
                //collect all AcDbText
                List<string> lstDbText = acDbText.Where(id => (Entity)tr.GetObject(id, OpenMode.ForRead) is DBText).Select(txtid => ((DBText)tr.GetObject(txtid, OpenMode.ForRead)).TextString.ToUpper()).ToList<string>();//
                if (_findtxt.Contains(','))
                {
                    foreach (string str in _findtxt.Split(','))
                        lstresultDbText.AddRange(lstDbText.Where(txt => txt.Contains(str.ToUpper())).Select(t => t.ToString()).ToList<string>());
                }
                else
                {
                    lstresultDbText = lstDbText.Where(txt => txt.Contains(_findtxt.ToUpper())).Select(t => t.ToString()).ToList<string>();
                }

                return (lstresultDbText.Count > 0 ? true : false);
                //List<string> contents = new List<string>();
                //foreach (var _text in acDbText)
                //{
                //    string content = string.Empty;
                //    Entity enttxt = (Entity)tr.GetObject(_text, OpenMode.ForWrite);
                //    if (enttxt is DBText)
                //    {

                //        DBText Text = ((DBText)enttxt);

                //        content = ((DBText)enttxt).TextString;
                //        layerNAME = Text.Layer;
                //        var textheight = Text.Height;
                //        var Location = Text.Position;
                //        var elename = _text.ObjectClass.Name;
                //        var meterial = Text.Material;

                //        if (content.Length >= 1)
                //        {

                //        }

                //    }
                //}
            }
            catch (Exception ex) { MessageBox.Show("", "" + ex.Message); }
            return false;
        }
        public static List<Curve> GetHatchBoundary(Hatch hatch)
        {
            int numberOfLoops = hatch.NumberOfLoops;
            var result = new List<Curve>(numberOfLoops);
            for (int i = 0; i < numberOfLoops; i++)
            {
                var loop = hatch.GetLoopAt(i);
                if (loop.IsPolyline)
                {
                    var bulges = loop.Polyline;
                    var pline = new Polyline(bulges.Count);
                    for (int j = 0; j < bulges.Count; j++)
                    {
                        var vertex = bulges[j];
                        pline.AddVertexAt(j, vertex.Vertex, vertex.Bulge, 0.0, 0.0);
                    }
                    pline.Elevation = hatch.Elevation;
                    pline.Normal = hatch.Normal;
                    result.Add(pline);
                }
                else
                {
                    var plane = hatch.GetPlane();
                    var xform = Matrix3d.PlaneToWorld(plane);
                    var curves = loop.Curves;
                    foreach (Curve2d curve in curves)
                    {
                        switch (curve)
                        {
                            case LineSegment2d lineSegment:
                                var line = new Line(
                                    new Point3d(lineSegment.StartPoint.X, lineSegment.StartPoint.Y, 0.0),
                                    new Point3d(lineSegment.EndPoint.X, lineSegment.EndPoint.Y, 0.0));
                                line.TransformBy(xform);
                                result.Add(line);
                                break;
                            case CircularArc2d circularArc:
                                double startAngle = circularArc.IsClockWise ? -circularArc.EndAngle : circularArc.StartAngle;
                                double endAngle = circularArc.IsClockWise ? -circularArc.StartAngle : circularArc.EndAngle;
                                var arc = new Arc(
                                    new Point3d(circularArc.Center.X, circularArc.Center.Y, 0.0),
                                    circularArc.Radius,
                                    circularArc.ReferenceVector.Angle + startAngle,
                                    circularArc.ReferenceVector.Angle + endAngle);
                                arc.TransformBy(xform);
                                result.Add(arc);
                                break;
                            case EllipticalArc2d ellipticalArc:
                                double ratio = ellipticalArc.MinorRadius / ellipticalArc.MajorRadius;
                                double startParam = ellipticalArc.IsClockWise ? -ellipticalArc.EndAngle : ellipticalArc.StartAngle;
                                double endParam = ellipticalArc.IsClockWise ? -ellipticalArc.StartAngle : ellipticalArc.EndAngle;
                                var ellipse = new Ellipse(
                                    new Point3d(ellipticalArc.Center.X, ellipticalArc.Center.Y, 0.0),
                                    Vector3d.ZAxis,
                                    new Vector3d(ellipticalArc.MajorAxis.X, ellipticalArc.MajorAxis.Y, 0.0) * ellipticalArc.MajorRadius,
                                    ratio,
                                    Math.Atan2(Math.Sin(startParam) * ellipticalArc.MinorRadius, Math.Cos(startParam) * ellipticalArc.MajorRadius),
                                    Math.Atan2(Math.Sin(endParam) * ellipticalArc.MinorRadius, Math.Cos(endParam) * ellipticalArc.MajorRadius));
                                ellipse.TransformBy(xform);
                                result.Add(ellipse);
                                break;
                            case NurbCurve2d nurbCurve:
                                var points = new Point3dCollection();
                                for (int j = 0; j < nurbCurve.NumControlPoints; j++)
                                {
                                    var pt = nurbCurve.GetControlPointAt(j);
                                    points.Add(new Point3d(pt.X, pt.Y, 0.0));
                                }
                                var knots = new DoubleCollection();
                                for (int k = 0; k < nurbCurve.NumKnots; k++)
                                {
                                    knots.Add(nurbCurve.GetKnotAt(k));
                                }
                                var weights = new DoubleCollection();
                                for (int l = 0; l < nurbCurve.NumWeights; l++)
                                {
                                    weights.Add(nurbCurve.GetWeightAt(l));
                                }
                                var spline = new Spline(nurbCurve.Degree, nurbCurve.IsRational, nurbCurve.IsClosed(), false, points, knots, weights, 0.0, 0.0);
                                spline.TransformBy(xform);
                                result.Add(spline);
                                break;
                            default:
                                break;
                        }
                    }
                }
            }
            return result;
        }

        public static List<Curve> GetHatchBoundary(AcadHatch hatch)
        {
            int numberOfLoops = hatch.NumberOfLoops;
            var result = new List<Curve>(numberOfLoops);
            for (int i = 0; i < numberOfLoops; i++)
            {
                object loop = null;
                hatch.GetLoopAt(i, out loop);


            }
            return result;
        }
        private static List<double> getlstLevels(List<ObjectId> acDbText, OpenCloseTransaction tr, string _layername)
        {
            List<double> lstLevels = new List<double>();
            try
            {

                //collect all AcDbText
                List<string> lstDbText = acDbText.Where(id => ((DBText)tr.GetObject(id, OpenMode.ForRead)).Layer.ToUpper().Contains(_layername.ToUpper())).Select(txtid => ((DBText)tr.GetObject(txtid, OpenMode.ForRead)).TextString).ToList<string>();//

                double x = 0;
                lstLevels = lstDbText.Where(str => double.TryParse(str.Replace("±", "").Replace("+", "").Replace("Peil", "0").Replace("-", ""), out x))
                               .Select(str => str.Contains("-") ? -x : x)
                               .ToList();
                //lstresultDbText = lstDbText.Where(txt => txt.Contains(_layername)).Select(t => t.ToString()).ToList<double>();

                lstLevels.Sort();
                return lstLevels;

            }
            catch (Exception ex) { MessageBox.Show("", "" + ex.Message); }
            return lstLevels;
        }
        private List<CAP.Level> LoadCollectionDataPDF()
        {
            List<CAP.Level> gLevels = new List<CAP.Level>();
            gLevels.Add(new CAP.Level()
            {
                LevelValue = 0,
                LevelNumbar = "0",
                LevelName = "Level 0"
            });
            gLevels.Add(new CAP.Level()
            {
                LevelValue = 3000,
                LevelNumbar = "1",
                LevelName = "Level 1"
            });
            gLevels.Add(new CAP.Level()
            {
                LevelValue = 6000,
                LevelNumbar = "2",
                LevelName = "Level 2"
            });
            gLevels.Add(new CAP.Level()
            {
                LevelValue = 9000,
                LevelNumbar = "3",
                LevelName = "Level 3"
            });
            gLevels.Add(new CAP.Level()
            {
                LevelValue = 12000,
                LevelNumbar = "4",
                LevelName = "Level 4"
            });
            gLevels.Add(new CAP.Level()
            {
                LevelValue = 15000,
                LevelNumbar = "5",
                LevelName = "Level 5"
            });

            return gLevels;
        }
        private List<CAP.Level> LoadCollectionDataDWG()
        {
            List<CAP.Level> gLevels = new List<CAP.Level>();
            gLevels.Add(new CAP.Level()
            {
                LevelValue = 0,
                LevelNumbar = "0",
                LevelName = "Level 0"
            });
            gLevels.Add(new CAP.Level()
            {
                LevelValue = 2880,
                LevelNumbar = "1",
                LevelName = "Level 1"
            });
            gLevels.Add(new CAP.Level()
            {
                LevelValue = 5760,
                LevelNumbar = "2",
                LevelName = "Level 2"
            });
            gLevels.Add(new CAP.Level()
            {
                LevelValue = 10550,
                LevelNumbar = "3",
                LevelName = "Level 3"
            });


            return gLevels;
        }
        #endregion Functions

        public class MyDriver
        {
            public static void CreateMyProfile()
            {
                bool isAutoCADRunning = Utilities.IsAutoCADRunning();
                if (isAutoCADRunning == false)
                    Utilities.StartAutoCADApp();
                Utilities.CreateProfile();
            }

            public static void NetloadMyApp(String dllPath)
            {
                bool isAutoCADRunning = Utilities.IsAutoCADRunning();
                if (isAutoCADRunning == false)
                    Utilities.StartAutoCADApp();
                Utilities.NetloadDll(dllPath);
            }
        }

        public class Utilities
        {
            [System.Runtime.InteropServices.DllImport("user32")]
            public static extern IntPtr GetWindowThreadProcessId(IntPtr hwnd, ref IntPtr lpdwProcessId);
            private static readonly string AutoCADProgname = "acad";
            private static readonly int sleepTime = 5000;
            private static readonly string AutoCADProgId = "AutoCAD.Application.24.1";
            private static AcadApplication app;
            public static AcadApplication App { get => app; set => app = value; }
            public static bool GetAcadInstance()
            {
                bool retValue = false;
                try
                {
                    if (Process.GetProcessesByName(AutoCADProgname).Count() == 0)
                    {
                        Type AppType = Type.GetTypeFromProgID(AutoCADProgId);
                        App = (AcadApplication)(Activator.CreateInstance(AppType));

                        //App.Visible = false;
                        //App.WindowState = Autodesk.AutoCAD.Interop.Common.AcWindowState.acMin;
                        retValue = true;
                    }
                    else
                    {
                        try
                        {

                            if (Marshal2.GetActiveObject(AutoCADProgId) != null)
                            {
                                App = (AcadApplication)Marshal2.GetActiveObject(AutoCADProgId);
                                Thread.Sleep(sleepTime);
                                //App.Visible = false;
                                //App.WindowState = Autodesk.AutoCAD.Interop.Common.AcWindowState.acMin;
                            }
                        }
                        catch (Exception ex)
                        {
                            KillAcad();
                            Type AppType = Type.GetTypeFromProgID(AutoCADProgId);
                            App = (AcadApplication)(Activator.CreateInstance(AppType));
                            Thread.Sleep(sleepTime);
                            // App.Visible = false;
                            //App.WindowState = Autodesk.AutoCAD.Interop.Common.AcWindowState.acMin;
                        }
                        retValue = true;
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                return retValue;
            }

            public static void KillAcad()
            {
                try
                {
                    foreach (Process process in Process.GetProcessesByName("acad"))
                    {
                        process.Kill();
                        System.Threading.Thread.Sleep(sleepTime);
                    }
                    foreach (Process process in Process.GetProcessesByName("senddmp"))
                    {
                        process.Kill();
                        System.Threading.Thread.Sleep(sleepTime);
                    }
                }
                catch (Exception ex) { }
            }
            public static void SendMessage(String message)
            {
                App.Application.ActiveDocument.SendCommand("(princ \"" + message + "\")(princ)" + Environment.NewLine);
            }

            public static bool IsAutoCADRunning()
            {
                bool isRunning = GetRunningAutoCADInstance();
                return isRunning;
            }

            public static bool ConfigureRunningAutoCADForUsage()
            {
                if (App == null)
                    return false;
                MessageFilter.Register();
                SetAutoCADWindowToNormal();
                return true;
            }

            public static bool StartAutoCADApp()
            {
                Type autocadType = System.Type.GetTypeFromCLSID(new Guid("AA46BA8A-9825-40FD-8493-0BA3C4D5CEB5"), true);
                object obj = System.Activator.CreateInstance(autocadType, true);
                AcadApplication appAcad = (AcadApplication)obj;
                App = appAcad;
                MessageFilter.Register();
                SetAutoCADWindowToNormal();
                return true;
            }


            public static bool NetloadDll(string dllPath)
            {
                if (!System.IO.File.Exists(dllPath))
                    throw new Exception("Dll does not exist: " + dllPath);
                App.ActiveDocument.SendCommand("(setvar \"secureload\" 0)" + Environment.NewLine);
                dllPath = dllPath.Replace(@"\", @"\\");
                App.ActiveDocument.SendCommand("(command \"_netload\" \"" + dllPath + "\")" + Environment.NewLine);
                return true;
            }


            public static bool CreateProfile()
            {
                if (App == null)
                    return false;
                bool profileExists = DoesProfileExist(App, yourProfileName);
                if (profileExists)
                {
                    SetYourProfileActive(App, yourProfileName);
                    AddTempFolderToTrustedPaths(App);
                }
                else
                {
                    CreateYourCustomProfile(App, yourProfileName);
                    AddTempFolderToTrustedPaths(App);
                }
                SetYourProfileActive(App, yourProfileName);
                return true;
            }


            public static bool SetAutoCADWindowToNormal()
            {
                if (App == null)
                    return false;
                App.WindowState = AcWindowState.acNorm;
                return true;
            }

            private static bool GetRunningAutoCADInstance()
            {
                Type autocadType = System.Type.GetTypeFromProgID(AutoCADProgId, true);
                AcadApplication appAcad;
                try
                {
                    object obj = Microsoft.VisualBasic.Interaction.GetObject(null, AutoCADProgId);
                    appAcad = (AcadApplication)obj;
                    App = appAcad;
                    return true;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                }
                return false;
            }

            public static readonly string yourProfileName = "myCustomProfile";

            private static void SetYourProfileActive(AcadApplication appAcad, string profileName)
            {
                AcadPreferencesProfiles profiles = appAcad.Preferences.Profiles;
                profiles.ActiveProfile = profileName;
            }

            private static void CreateYourCustomProfile(AcadApplication appAcad, string profileName)
            {
                AcadPreferencesProfiles profiles = appAcad.Preferences.Profiles;
                profiles.CopyProfile(profiles.ActiveProfile, profileName);
                profiles.ActiveProfile = profileName;
            }

            private static bool DoesProfileExist(AcadApplication appAcad, string profileName)
            {
                AcadPreferencesProfiles profiles = appAcad.Preferences.Profiles;
                object pNames = null;
                profiles.GetAllProfileNames(out pNames);
                string[] profileNames = (string[])pNames;
                foreach (string name in profileNames)
                {
                    if (name.Equals(profileName))
                        return true;
                }
                return false;
            }

            private static void AddTempFolderToTrustedPaths(AcadApplication appAcad)
            {
                string trustedPathsString = System.Convert.ToString(appAcad.ActiveDocument.GetVariable("TRUSTEDPATHS"));
                string tempDirectory = System.IO.Path.GetTempPath();
                List<string> newPaths = new List<string>() { tempDirectory };
                if (!trustedPathsString.Contains(tempDirectory))
                    AddTrustedPaths(appAcad, newPaths);
            }

            private static void AddTrustedPaths(AcadApplication appAcad, List<string> newPaths)
            {
                string trustedPathsString = System.Convert.ToString(appAcad.ActiveDocument.GetVariable("TRUSTEDPATHS"));
                List<string> oldPaths = new List<string>();
                oldPaths = trustedPathsString.Split(System.Convert.ToChar(";")).ToList();
                string newTrustedPathsString = trustedPathsString;
                foreach (string newPath in newPaths)
                {
                    bool pathAlreadyExists = trustedPathsString.Contains(newPath);
                    if (!pathAlreadyExists)
                        newTrustedPathsString = newPath + ";" + newTrustedPathsString;
                }
                appAcad.ActiveDocument.SetVariable("TRUSTEDPATHS", newTrustedPathsString);
            }
        }

        public class MessageFilter : IOleMessageFilter
        {
            [DllImport("Ole32.dll")]
            private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, ref IOleMessageFilter oldFilter);

            public static void Register()
            {
                IOleMessageFilter newFilter = new MessageFilter();
                IOleMessageFilter oldFilter = null;
                CoRegisterMessageFilter(newFilter, ref oldFilter);
            }
            public static void Revoke()
            {
                IOleMessageFilter oldFilter = null;
                CoRegisterMessageFilter(null, ref oldFilter);
            }

            public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
            {
                return 0;
            }

            public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
            {
                if (dwRejectType == 2)
                    // flag = SERVERCALL_RETRYLATER.

                    // Retry the thread call immediately if return >=0 & 
                    // <100.
                    return 99;
                // Too busy; cancel call.
                return -1;
            }

            public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
            {
                return 2;
            }
        }

        [ComImport()]
        [Guid("00000016-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IOleMessageFilter
        {
            [PreserveSig]
            int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
            [PreserveSig]
            int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
            [PreserveSig]
            int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
        }

        private void dgv_levels_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            e.Cancel = true;
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {

            using (var fbd = new FolderBrowserDialog())
            {
                fbd.SelectedPath = Properties.Settings.Default.catchEcapPath;
                DialogResult result = fbd.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    Properties.Settings.Default.catchEcapPath = fbd.SelectedPath;

                    txtBrowse.Text = fbd.SelectedPath;
                    Properties.Settings.Default.Save();
                    lstFiles = Directory.GetFiles(fbd.SelectedPath, "*.pdf*").ToList();
                    lstFiles.AddRange(Directory.GetFiles(fbd.SelectedPath, "*.dwg*").ToList());
                    ECAP.EcapLogin.lstFiles = lstFiles.ToArray();
                }
            }

        }

        private void btnScan_Click(object sender, RoutedEventArgs e)
        {
            btnScan.IsEnabled = false;
            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;

            countHatchs = 0;
            lstLevelAreas = new List<string>();
            lstLevels = new List<double>();
            levelHatchs = new SortedList<string, List<Curve>>();
            wallPolyline = new SortedList<string, List<Curve>>();

            List<string> _fileslist = new List<string>();
            if (txtBrowse.Text.Length <= 0)
                using (var fbd = new FolderBrowserDialog())
                {
                    fbd.SelectedPath = Properties.Settings.Default.catchEcapPath;
                    DialogResult result = fbd.ShowDialog();
                    if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {
                        Properties.Settings.Default.catchEcapPath = fbd.SelectedPath;

                        txtBrowse.Text = fbd.SelectedPath;
                        Properties.Settings.Default.Save();
                        lstFiles = Directory.GetFiles(fbd.SelectedPath, "*.pdf*").ToList();
                        lstFiles.AddRange(Directory.GetFiles(fbd.SelectedPath, "*.dwg*").ToList());

                    }
                }
            ECAP.EcapLogin.lstFiles = lstFiles.ToArray();
            filecount = lstFiles.Count();
            proBar.Maximum = filecount;
            proBar.Minimum = 0;
            proBar.Value = 0;
            System.Windows.Forms.Application.DoEvents();
            int countSection = 0;
            int countElevation = 0;
            int countFloor = 0;


            if (txtBrowse_Output.Text.Length <= 0)
            {
                SaveFileDialog openFileDialog1 = new SaveFileDialog
                {
                    InitialDirectory = Properties.Settings.Default.catchIcapPath,
                    Title = "Browse Cap Files",
                    CheckFileExists = false,
                    CheckPathExists = true,
                    DefaultExt = "cap",
                    Filter = "Cap files (*.cap)|*.cap",
                    FilterIndex = 2,
                    RestoreDirectory = true,

                };

                if (openFileDialog1.ShowDialog() == true)
                {
                    ECAP.EcapLogin.capfilename = openFileDialog1.FileName;
                    filename = openFileDialog1.FileName;
                    Properties.Settings.Default.catchIcapPath = filename;
                    txtBrowse_Output.Text = filename;
                    Properties.Settings.Default.Save();
                }
            }
            ECAP.EcapLogin.capfilename = txtBrowse_Output.Text;
            filename = Path.GetTempPath() + Path.DirectorySeparatorChar.ToString() + "temp.cap";

            //System.Xml.XmlWriterSettings settings = new System.Xml.XmlWriterSettings();
            //settings.Indent = true;
            //settings.IndentChars = "    ";
            //settings.OmitXmlDeclaration = false;
            //settings.Encoding = System.Text.Encoding.UTF8;


            //using (System.Xml.XmlWriter writer = System.Xml.XmlWriter.Create(filename, settings))
            //{
            try
            {
                //writer.WriteStartDocument();
                //writer.WriteStartElement("Export");
                //writer.WriteStartElement("CAD");
                foreach (string _file in lstFiles)
                {
                    proBar.Value++;
                    System.Windows.Forms.Application.DoEvents();
                    try
                    {
                        isSection = false;
                        isElevation = false;
                        isFloor = false;

                        currentFilename = System.IO.Path.GetFileNameWithoutExtension(_file).ToString();
                        if (_file != null && _file != "")
                        {
                            //writer.WriteStartElement("File");
                            //writer.WriteAttributeString("FileName", System.IO.Path.GetFileNameWithoutExtension(_file).ToString());
                            //writer.WriteAttributeString("FilePath", _file.ToString());

                            if (IsFileWritable(_file) & _file.ToLower().EndsWith("dwg"))
                            {
                                #region CAD
                                //  AcadDocument _acadDoc = Utilities.App.Documents.Open(_file);
                                bool isAcad = Utilities.GetAcadInstance();
                                //MyDriver.CreateMyProfile();
                                // Utilities.SendMessage("Profile created:" + Utilities.yourProfileName);
                                //MyDriver.NetloadMyApp(@"C:\YourPathFolderPath\custom.Dll");

                                if (isAcad)
                                {
                                    //Utilities.SendMessage("AutoCAD started from WPF");
                                    // Display the application and return the name and version
                                    //Utilities.App.Visible = false;
                                    //AcadApp.Documents.Open(@"C:\Templates\TEST_BASE.dwt");
                                    var AcadDoc = Utilities.App.Documents.Open(_file);
                                    try
                                    {
                                        foreach (AcadEntity item in AcadDoc.ModelSpace)
                                        {
                                            if (item.ObjectName == "AcDbHatch")
                                            {
                                                AcadHatch hatch = ((AcadHatch)item);
                                                var elename = hatch.ObjectName;
                                                var type = hatch.PatternType;
                                                //var Maximun = hatch.GetBoundingBox().MaxPoint;
                                                //var Minimum = hatch.GeometricExtents.MinPoint;
                                                var spacing = hatch.PatternSpace;
                                                var Hatchweight = hatch.Lineweight;
                                                var meterial = hatch.Material;
                                                var area = hatch.Area;
                                                string areastr = Math.Round(area, 0).ToString();

                                                //List<Curve> curves = GetHatchBoundary(hatch);
                                                layerNAME = hatch.Layer;

                                                //countHatchs++;
                                                //if (isFloor && areastr.Length > 7 && !lstLevelAreas.Contains(areastr + currentFilename.ToLower().Replace(" ", "")))
                                                //{
                                                //    lstLevelAreas.Add(areastr + currentFilename.ToLower().Replace(" ", ""));
                                                //    levelHatchs.Add("Hatch" + countHatchs.ToString() + ";" + currentFilename.ToLower().Replace(" ", ""), curves);
                                                //}
                                                //pattendtype
                                                //writer.WriteStartElement("Hatch");
                                                //writer.WriteAttributeString("ElementID", "" + hatch.ToString());
                                                //writer.WriteAttributeString("ElementName", "" + elename.ToString());
                                                //writer.WriteAttributeString("LayerName", "" + layerNAME.ToString());
                                                //writer.WriteAttributeString("Type", "" + type.ToString());
                                                //writer.WriteAttributeString("Space", "" + spacing.ToString());
                                                //writer.WriteAttributeString("HatchWeight", "" + Hatchweight.ToString());
                                                //writer.WriteAttributeString("Material", "" + meterial.ToString());
                                                //writer.WriteAttributeString("Area", "" + area.ToString());
                                                //writer.WriteElementString("Maximun", "" + Maximun.ToString());
                                                //writer.WriteElementString("Minimum", "" + Minimum.ToString());
                                                //writer.WriteStartElement("Boundary");
                                                int xcount = 0;
                                                for (var curve = 0; curve <= hatch.NumberOfLoops; curve++)
                                                {
                                                    object acadObject;
                                                    hatch.GetLoopAt(curve, out acadObject);
                                                    if (acadObject != null)
                                                    {
                                                        var objs = (object[])acadObject;
                                                        var _acad = objs[0];
                                                        AcadEntity acadEntity = (AcadEntity)_acad;
                                                        if (acadEntity.ObjectName == "AcDbPolyline")
                                                        {
                                                            var pline = (Acad3DPolyline)acadEntity;
                                                        

                                                        }



                                                        //writer.WriteElementString("cPoint", "" + curve.StartPoint.ToString());
                                                        xcount++;
                                                    }
                                                    //if (xcount == curves.Count)
                                                    //writer.WriteElementString("cPoint", "" + curve.EndPoint.ToString());
                                                }
                                                //writer.WriteEndElement();
                                                //writer.WriteEndElement();

                                            }
                                        }
                                    }
                                    catch(Exception ex) { MessageBox.Show(ex.Message, "CapInfo"); }
                                    AcadDoc.Close();
                                    ////////AcadDoc.ModelSpace.GetEnumerator().MoveNext();
                                    ////////Autodesk.AutoCAD.ApplicationServices.Document cadDoc = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.FromAcadDocument(AcadDoc);
                                    ////////(command "_netload" "C:/CAD/C3D/2021/Basics/CivYam.dll")
                                    //////AcadDoc.SendCommand(@"(command ""_netload"" """ + dlllocation.Replace(@"\", @"\\") + @""") ");
                                    ////////(comman "_scan")
                                    //////AcadDoc.SendCommand(@"(command ""CAPSCAN"") ");
                                }
                                #endregion CAD
                            }

                            else if (IsFileWritable(_file) & _file.ToLower().EndsWith("pdf"))
                            {
                                #region PDF
                                //using (var fs = File.OpenRead(@"C:\Users\M S P\Documents\CAP\2D to 3D\Development\Sample\Sourcefiles deken frankenstraat\Floorplan 1e.pdf"))
                                //{
                                //    //Load Sample PDF document  
                                //    GcPdfDocument doc = new GcPdfDocument();
                                //    doc.Load(fs);

                                //    //Get File Attachment using FileAttachmentAnnotation   
                                //    //Get first page from document  
                                //    Page page = doc.Pages[0];

                                //    //Get the annotation collection from pages  
                                //    PageContentStreamCollection contents = page.ContentStreams;
                                //    proBar.Maximum = contents.Count;
                                //    proBar.Minimum = 0;
                                //    proBar.Value = 0;
                                //    //Iterates the annotations  
                                //    foreach (ContentStream contentStream in contents)
                                //    {
                                //        proBar.Value++;
                                //        //Check for the attachment annotation  
                                //        if (contentStream.HasContent())
                                //        {

                                //            //Extracts the attachment and saves it to the disk  
                                //            //string path = @"C:\Users\M S P\Documents\CAP\2D to 3D\Development\Sample\Sourcefiles deken frankenstraat\";
                                //            //FileStream stream = new FileStream(path, FileMode.Create);
                                //            Stream stream = contentStream.GetStream();

                                //            stream.Dispose();
                                //        }
                                //    }


                                //    //Extract All text from PDF document  
                                //    var texttest = doc.GetText();
                                //    //Display the results:  
                                //    Console.WriteLine("PDF Text: \n \n" + texttest);

                                //    //Extract specific page text:  
                                //    var pageText = doc.Pages[0].GetText();
                                //    //Display the results:  
                                //    Console.WriteLine("PDF Page Text: \n" + pageText);
                                //}
                                #endregion PDF
                            }

                            // writer.WriteEndElement();
                        }
                        Mouse.OverrideCursor = System.Windows.Input.Cursors.Arrow;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "CapInfo");
                    }
                }

                // writer.WriteEndElement();
                //Microsoft.WindowsAPICodePack.Dialogs.TaskDialog.Show("CapInfo", "Completed");


            }
            finally
            {
                btnScan.IsEnabled = true;
                Mouse.OverrideCursor = System.Windows.Input.Cursors.Arrow;
                MessageBox.Show(@"Total " + filecount + " *.dwg files are deducted! in this we have below : \n\n"
                        + "Section " + countSection + "*.dwg files.\n"
                        + "Elevation " + countElevation + "*.dwg files.\n"
                        + "Plan " + countFloor + "*.dwg files.\n" + Environment.NewLine + @"Total " + filecount + " *.pdf files are deducted! in this we have below : \n\n"
                        + "Section " + countSection + "*.pdf files.\n"
                        + "Elevation " + countElevation + "*.pdf files.\n"
                        + "Plan " + countFloor + "*.pdf files.\n", "CapInfo");
            }

            //}

        }

        private void btnBrowse_output_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog openFileDialog1 = new SaveFileDialog
            {
                InitialDirectory = Properties.Settings.Default.catchIcapPath,
                Title = "Browse Cap Files",
                CheckFileExists = false,
                CheckPathExists = true,
                DefaultExt = "cap",
                Filter = "Cap files (*.cap)|*.cap",
                FilterIndex = 2,
                RestoreDirectory = true,

            };

            if (openFileDialog1.ShowDialog() == true)
            {
                ECAP.EcapLogin.capfilename = openFileDialog1.FileName;
                filename = openFileDialog1.FileName;
                Properties.Settings.Default.catchIcapPath = filename;
                txtBrowse_Output.Text = filename;
                Properties.Settings.Default.Save();

            }
        }

        private void btnCap_Click(object sender, RoutedEventArgs e)
        {
            //filecount = lstFiles.Count();

            //int countSection = 0;
            //int countElevation = 0;
            //int countFloor = 0;
            //try
            //{
            //    XmlDocument xmlDocument = new XmlDocument();
            //    xmlDocument.LoadXml(filename);

            //    System.Xml.XmlWriterSettings settings = new System.Xml.XmlWriterSettings();
            //    settings.Indent = true;
            //    settings.IndentChars = "    ";
            //    settings.OmitXmlDeclaration = false;
            //    settings.Encoding = System.Text.Encoding.UTF8;


            //    using (System.Xml.XmlWriter writer = System.Xml.XmlWriter.Create(txtBrowse_Copy.Text, settings))
            //    {
            //        xmlDocument.WriteTo(writer);

            //        writer.WriteEndElement();
            //        writer.WriteStartElement("CAP");
            //        //Wallhatchdetails(AcDbHatch, writer, tr);
            //        writer.WriteStartElement("Levels");
            //        SortedList<string, string> listleveldetails = new SortedList<string, string>();
            //        int countLevel = 0;
            //        //frmLevel frmLevel = new frmLevel(lstLevels, levelHatchs);
            //        //frmLevel.ShowDialog();
            //        if (lstLevels.Count > 0)
            //        {
            //            foreach (double levelVal in lstLevels)
            //            {
            //                if (levelVal >= 0)
            //                {

            //                    writer.WriteStartElement("Level" + countLevel);
            //                    string levelname = countLevel == 0 ? "Ground Floor" : (countLevel == 1 ? "1st Floor" : (countLevel == 2 ? "2nd Floor" : "3rd Floor"));
            //                    writer.WriteElementString("Level_Name", levelname);
            //                    writer.WriteElementString("Level_Number", countLevel.ToString());
            //                    writer.WriteElementString("Level_Value", levelVal.ToString());
            //                    writer.WriteEndElement();
            //                    if (!listleveldetails.ContainsKey(levelname.ToLower().Replace(" ", "")))
            //                        listleveldetails.Add(levelname.ToLower().Replace(" ", ""), levelVal.ToString());
            //                    countLevel++;
            //                }
            //            }
            //        }
            //        writer.WriteEndElement();

            //        writer.WriteStartElement("Floors");

            //        foreach (string levelkey in levelHatchs.Keys)
            //        {
            //            List<Curve> curves = levelHatchs[levelkey];
            //            writer.WriteStartElement(levelkey.Split(';')[0].ToString());
            //            string lvname = levelkey.Split(';')[1].ToString();
            //            writer.WriteElementString("Level_Name", lvname);

            //            if (listleveldetails.ContainsKey(lvname))
            //            {
            //                writer.WriteElementString("Level_Number", (lvname == "groundfloor" ? "0" : (lvname == "1stfloor" ? "1" : (lvname == "2ndfloor" ? "2" : "3"))));
            //                writer.WriteElementString("Level_Value", listleveldetails[lvname]);
            //            }
            //            int xcount = 0;
            //            double zValue = 0;
            //            if (listleveldetails.ContainsKey(lvname))
            //                double.TryParse(listleveldetails[lvname], out zValue);
            //            foreach (var curve in curves)
            //            {
            //                writer.WriteElementString("cPoint", "" + new Point3d(curve.StartPoint.X, curve.StartPoint.Y, zValue).ToString());
            //                xcount++;
            //                if (xcount == curves.Count)
            //                    writer.WriteElementString("cPoint", "" + new Point3d(curve.EndPoint.X, curve.EndPoint.Y, zValue).ToString());
            //            }
            //            writer.WriteEndElement();
            //        }
            //        writer.WriteEndElement();
            //        writer.WriteStartElement("Walls");

            //        foreach (string wallkey in wallPolyline.Keys)
            //        {
            //            List<Curve> curves = wallPolyline[wallkey];
            //            writer.WriteStartElement(wallkey.Split(';')[0].ToString());
            //            string lvname = wallkey.Split(';')[1].ToString();
            //            writer.WriteElementString("Level_Name", lvname);

            //            if (listleveldetails.ContainsKey(lvname))
            //            {
            //                writer.WriteElementString("Level_Number", (lvname == "groundfloor" ? "0" : (lvname == "1stfloor" ? "1" : (lvname == "2ndfloor" ? "2" : "3"))));
            //                writer.WriteElementString("Level_Value", listleveldetails[lvname]);
            //            }
            //            int xcount = 0;
            //            double zValue = 0;
            //            if (listleveldetails.ContainsKey(lvname))
            //                double.TryParse(listleveldetails[lvname], out zValue);
            //            foreach (var curve in curves)
            //            {
            //                writer.WriteElementString("cPoint", "" + new Point3d(curve.StartPoint.X, curve.StartPoint.Y, zValue).ToString());
            //                xcount++;
            //                if (xcount == curves.Count)
            //                    writer.WriteElementString("cPoint", "" + new Point3d(curve.EndPoint.X, curve.EndPoint.Y, zValue).ToString());
            //            }
            //            writer.WriteEndElement();
            //        }
            //        writer.WriteEndElement();

            //        writer.Flush();
            //        writer.Close();
            //    }

            //    MessageBox.Show(@"Total " + filecount + " *.dwg files are deducted! in this we have below : \n\n"
            //        + "Section " + countSection + "*.dwg files.\n"
            //        + "Elevation " + countElevation + "*.dwg files.\n"
            //        + "Plan " + countFloor + "*.dwg files.\n", "CapInfo");
            //    //Microsoft.WindowsAPICodePack.Dialogs.TaskDialog.Show("CapInfo", "Completed");
            //}
            //catch (Exception ex) { MessageBox.Show(ex.Message, "CapInfo"); }

            System.IO.File.Copy(_outputcapfilename, txtBrowse_Output.Text, true);

            XElement rootElement = XElement.Load(txtBrowse_Output.Text);

            XElement xLevels = rootElement.Element("CAP").Element("Levels");

            SortedList<string, string> listleveldetails = new SortedList<string, string>();
            int countLevel = 0;
            foreach (var levelVItem in CAP.Levels)
            {
                //if (levelVItem is not gLevel) continue;
                XElement xLevel = xLevels.Element((levelVItem).LevelName.Replace(" ", ""));
                //writer.WriteStartElement("Level" + countLevel);
                string levelname = countLevel == 0 ? "Ground Floor" : (countLevel == 1 ? "1st Floor" : (countLevel == 2 ? "2nd Floor" : "3rd Floor"));
                if (xLevel == null)
                {
                    xLevels.Add(new XElement((levelVItem).LevelName.Replace(" ", "")));
                    xLevel = xLevels.Element((levelVItem).LevelName.Replace(" ", ""));
                }
                xLevel.SetElementValue("Level_Name", (levelVItem).LevelName);
                xLevel.SetElementValue("Level_Number", (levelVItem).LevelNumbar.ToString());
                xLevel.SetElementValue("Level_Value", (levelVItem).LevelValue.ToString());


                countLevel++;

            }
            rootElement.Save(txtBrowse_Output.Text);
            MessageBox.Show(@"The file {" + System.IO.Path.GetFileName(txtBrowse_Output.Text) + "} has exported\n", "CapInfo");


        }

        private void DataGridCell_Selected(object sender, RoutedEventArgs e)
        {
            // Lookup for the source to be DataGridCell
            if (e.OriginalSource.GetType() == typeof(DataGridCell))
            {
                // Starts the Edit on the row;
                DataGrid grd = (DataGrid)sender;
                grd.BeginEdit(e);
            }
        }

        private void btnAI_Click(object sender, RoutedEventArgs e)
        {
            dgv_levels.EnableColumnVirtualization = true;
            var _levels = new List<CAP.Level>();
            //if (dgv_levels.ItemsSource == null)
            //{
            _levels = txtBrowse.Text.ToLower().Contains("deken") ? LoadCollectionDataPDF() : LoadCollectionDataDWG();
            dgv_levels.ItemsSource = _levels;
            //}
            //else
            //{
            //  //var  _levelsnew = dgv_levels.ItemsSource;
            //  //  _levels = _levelsnew.ToList<CAP.Level>();
            //  foreach(DataGridRow row in dgv_levels.ItemsSource)
            //    {
            //        CAP.Level level = new CAP.Level();
            //        level.LevelName = row.;

            //        _levels.Add()
            //    }
            //}
            int currentlevel = 0;
            tab_Floors.Items.Clear();
            tab_Walls.Items.Clear();
            XElement xElement = XElement.Load(_outputcapfilename.Replace("dkn", "flwa"));
            foreach (var level in _levels)
            {
                TabItem tabItemlevel = new TabItem();
                tabItemlevel.Header = level.LevelName;
                tab_Floors.Items.Add(tabItemlevel);
                TabItem _floorview = (TabItem)tab_Floors.Items.GetItemAt(currentlevel);
                Canvas CanvasView = new Canvas();
                CanvasView.RenderTransform = RenderTransform;
                System.Windows.Shapes.Polyline pline = new System.Windows.Shapes.Polyline();
                PointCollection pointCollection = new PointCollection();
                CanvasView.Children.Clear();


                if (pline != null) // pline = PolyLine
                {
                    pline.Stroke = new SolidColorBrush(Color.FromRgb(60, 125, 200));
                    pline.StrokeThickness = 3;

                    //    < cPoint > (0, 11030, 5760) </ cPoint >
                    //< cPoint > (6800, 11030, 5760) </ cPoint >
                    //< cPoint > (6800, 0, 5760) </ cPoint >
                    //< cPoint > (0, 0, 5760) </ cPoint >
                    //< cPoint > (0, 11030, 5760) </ cPoint >
                    pointCollection = new PointCollection();
                    foreach (XElement xE in xElement.Element("CAP").Element("Floors").Elements())
                    {
                        foreach (XElement xEle in xE.Elements("cPoint"))
                        {
                            if (level.LevelNumbar != xE.Element("Level_Number").Value) continue;
                            double xValue = 0;
                            double.TryParse(xEle.Value.Replace("(", "").Replace(")", "").Split(',')[0].ToString(), out xValue);
                            double yValue = 0;
                            double.TryParse(xEle.Value.Replace("(", "").Replace(")", "").Split(',')[1].ToString(), out yValue);
                            xValue = xValue / slidervalue;
                            yValue = yValue / slidervalue;
                            pointCollection.Add(new Point(xValue, yValue));
                        }
                    }

                    pline.Points = pointCollection;

                    //slider.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                    //slider.VerticalAlignment = VerticalAlignment.Bottom;
                    //slider.Margin= new Thickness(3, 0, 0, 3);
                    CanvasView.Children.Add(pline);

                }
                //CanvasView.Width = _floorview.Width;
                //CanvasView.Height = _floorview.Height;
                // CanvasView.Background= new SolidColorBrush(Color.FromRgb(60, 125, 200)); 
                Drawer drawer = new Drawer(CanvasView);
                drawer.ContinuousDraw = true;
                drawer.DrawTool = Tool.Selection;
                System.Windows.Controls.ComboBox combobox = new System.Windows.Controls.ComboBox();
                combobox.Items.Add("System Family: Floor");
                combobox.SelectedIndex = 0;
                combobox.Width = 200;
                combobox.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                //CanvasView.Children.Add(combobox);


                //CanvasView.Children.Add(slider);
                StackPanel stackPanel = new StackPanel();
                stackPanel.Children.Add(combobox);
                stackPanel.Children.Add(CanvasView);
                _floorview.Content = stackPanel;


                TabItem tabItemwall = new TabItem();
                tabItemwall.Header = level.LevelName;
                tab_Walls.Items.Add(tabItemwall);
                TabItem _wallview = (TabItem)tab_Walls.Items.GetItemAt(currentlevel);
                Canvas CanvaswView = new Canvas();
                pline = new System.Windows.Shapes.Polyline();
                pointCollection = new PointCollection();
                CanvaswView.Children.Clear();
                if (pline != null) // pline = PolyLine
                {
                    pline.Stroke = new SolidColorBrush(Color.FromRgb(60, 125, 200));
                    pline.StrokeThickness = 3;

                    pointCollection = new PointCollection();
                    foreach (XElement xE in xElement.Element("CAP").Element("Walls").Elements())
                    {
                        foreach (XElement xEle in xE.Elements("cPoint"))
                        {
                            if (level.LevelNumbar != xE.Element("Level_Number").Value) continue;
                            double xValue = 0;
                            double.TryParse(xEle.Value.Replace("(", "").Replace(")", "").Split(',')[0].ToString(), out xValue);
                            double yValue = 0;
                            double.TryParse(xEle.Value.Replace("(", "").Replace(")", "").Split(',')[1].ToString(), out yValue);
                            xValue = xValue / slidervalue;
                            yValue = yValue / slidervalue;
                            pointCollection.Add(new Point(xValue, yValue));
                        }
                    }
                    pline.Points = pointCollection;
                    CanvaswView.Children.Add(pline);
                }
                //CanvasView.Width = _floorview.Width;
                //CanvasView.Height = _floorview.Height;
                // CanvasView.Background= new SolidColorBrush(Color.FromRgb(60, 125, 200)); 
                drawer = new Drawer(CanvasView);
                drawer.ContinuousDraw = true;
                drawer.DrawTool = Tool.Selection;
                combobox = new System.Windows.Controls.ComboBox();
                combobox.Items.Add("System Family: Basic Wall");
                combobox.Items.Add("System Family: Brick Wall");
                combobox.Items.Add("System Family: Curtain Wall");
                combobox.Items.Add("System Family: Stacked Wall");
                combobox.SelectedIndex = 0;
                CanvaswView.Children.Add(combobox);
                _wallview.Content = CanvaswView;

                currentlevel++;
            }
        }

        private void Slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (e.NewValue != 10)
            {
                slidervalue = e.NewValue;
                btnAI_Click(null, null);
            }
        }
    }
}
