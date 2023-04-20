using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.IO;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.PlottingServices;
using System.Windows.Forms;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.Geometry;
using System.Xml;
using Exception = Autodesk.AutoCAD.Runtime.Exception;
using System.Security.Cryptography;
using Autodesk.Windows;
using System.Windows.Media.Imaging;
using System.Reflection;
using Microsoft.WindowsAPICodePack.Dialogs;
using OpenFileDialog = Autodesk.AutoCAD.Windows.OpenFileDialog;
using ECAP.Properties;
using System.Runtime;
using Autodesk.AutoCAD.Ribbon;
using System.Windows.Shapes;
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;
using Path = System.IO.Path;
using Line = Autodesk.AutoCAD.DatabaseServices.Line;
using Ellipse = Autodesk.AutoCAD.DatabaseServices.Ellipse;
using System.Drawing;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Input;

namespace ECAP
{
    public class EcapLogin
    {
        public static string layerNAME;
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
        public static string[] lstFiles = null;
        private static int filecount;
        public static string capfilename = string.Empty;

        [CommandMethod("CAP")]
        public static void CAP()
        {
            string Capimage = Path.Combine(Path.GetFullPath(System.Reflection.Assembly.GetExecutingAssembly().Location)
                .Remove(Path.GetFullPath(System.Reflection.Assembly.GetExecutingAssembly().Location).Count()
                - Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location).Count()), @"Resource\logoBOT.jpg");
            #region tab
            RibbonControl ribbonControl = ComponentManager.Ribbon;

            // Creation of Tab
            RibbonTab Tab = null; ;
            if (ribbonControl.Tabs.Count != 0)
            {
                foreach (RibbonTab ribbonTab in ribbonControl.Tabs)
                {
                    if (ribbonTab.Title == "Cap".ToUpper())
                    {
                        Tab = ribbonTab;
                    }
                }
            }
            if (Tab == null)
            {
                Tab = new RibbonTab();
                Tab.Title = "Cap".ToUpper();
                Tab.Id = "Capinfo";
                ribbonControl.Tabs.Add(Tab);
            }
            #endregion TAB
            RibbonPanelSource srcPanel = new RibbonPanelSource();
            RibbonPanel ribbonpanel = null;
            if (Tab.Panels != null)
            {
                foreach (RibbonPanel panel in Tab.Panels)
                {
                    if (panel.UID == "cappanel".ToUpper())
                    {
                        ribbonpanel = panel;
                    }
                }
            }
            if (ribbonpanel == null)
            {
                ribbonpanel = new RibbonPanel();
                ribbonpanel.UID = "cappanel".ToUpper();
                srcPanel.Title = "Cap Info";
                ribbonpanel.Source = srcPanel;
                Tab.Panels.Add(ribbonpanel);
            }
            Autodesk.Windows.RibbonButton btnPower = new Autodesk.Windows.RibbonButton();
            btnPower.Orientation = (System.Windows.Controls.Orientation)Orientation.Vertical;
            btnPower.Text = "EXPORT CAP";
            btnPower.Id = "btnCap";
            btnPower.ShowText = true;

            btnPower.LargeImage = getBitmap(Capimage, 32, 32);
            btnPower.Size = RibbonItemSize.Large;
            btnPower.CommandHandler = new RibbonButton();
            srcPanel.Items.Add(btnPower);
            Autodesk.Windows.RibbonButton btnEcap = new Autodesk.Windows.RibbonButton();
            btnEcap.Orientation = (System.Windows.Controls.Orientation)Orientation.Vertical;
            btnEcap.Text = "E-CAP";
            btnEcap.Id = "btnEcap";
            btnEcap.ShowText = true;

            btnEcap.LargeImage = getBitmap(Capimage, 32, 32);
            btnEcap.Size = RibbonItemSize.Large;
            btnEcap.CommandHandler = new RibbonButton();
            //srcPanel.Items.Add(new RibbonSeparator());
            srcPanel.Items.Add(btnEcap);
        }

        [CommandMethod("CAPSCAN")]
        public static void CAPScan()
        {

            countHatchs = 0;
            lstLevelAreas = new List<string>();
            lstLevels = new List<double>();
            levelHatchs = new SortedList<string, List<Curve>>();
            wallPolyline = new SortedList<string, List<Curve>>();

            List<string> _fileslist = new List<string>();
        
            //lstFiles = Directory.GetFiles(Properties.Settings.Default.catchEcapPath, "*.dwg*");
            filecount = lstFiles.Count();
            int countSection = 0;
            int countElevation = 0;
            int countFloor = 0;


           // capfilename = Path.GetTempPath() + Path.DirectorySeparatorChar.ToString() + "temp.cap";
            try
            {
                System.Xml.XmlWriterSettings settings = new System.Xml.XmlWriterSettings();
                settings.Indent = true;
                settings.IndentChars = "    ";
                settings.OmitXmlDeclaration = false;
                settings.Encoding = System.Text.Encoding.UTF8;


                using (System.Xml.XmlWriter writer = System.Xml.XmlWriter.Create(capfilename, settings))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("Export");
                    writer.WriteStartElement("CAD");

                    foreach (string _file in lstFiles)
                    {
                        isSection = false;
                        isElevation = false;
                        isFloor = false;

                        currentFilename = System.IO.Path.GetFileNameWithoutExtension(_file).ToString();
                        if (_file != null && _file != "")
                        {
                            writer.WriteStartElement("File");
                            writer.WriteAttributeString("FileName", System.IO.Path.GetFileNameWithoutExtension(_file).ToString());
                            writer.WriteAttributeString("FilePath", _file.ToString());


                            if (IsFileWritable(_file))
                                // using (var acDOCLOCK = doc.LockDocument())
                                using (Database db = new Database(false, false))
                                {
                                    Autodesk.AutoCAD.Windows.Window window = Autodesk.AutoCAD.ApplicationServices.Core.Application.MainWindow;
                                    //var dbfrmAcad = Database.GetAllDatabases();// FromAcadDatabase(AcadDB);
                                    db.ReadDwgFile(_file, FileOpenMode.OpenTryForReadShare, true, "");

                                    db.CloseInput(true);
                                    //var acadDB = db.AcadDatabase;
                                    //HostApplicationServices.WorkingDatabase = dbfrmAcad;
                                    // var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.GetDocument(db);//
                                    // Autodesk.AutoCAD.ApplicationServices.Document acDoc = doc.Application.Documents..GetDocument(db);
                                    using (var tr = db.TransactionManager.StartOpenCloseTransaction())// StartOpenCloseTransaction())
                                    {
                                        //reading all details from the model space
                                        var modelSpace = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                                        var brclass = RXObject.GetClass(typeof(BlockReference));
                                        //collecting all blocks in the Model 
                                        var blocks = modelSpace.Cast<ObjectId>().Where(id => id.ObjectClass == brclass).Select(id => (BlockReference)tr.GetObject(id, OpenMode.ForRead))
                                            .GroupBy(br => ((BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead)).Name);

                                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);


                                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                                        // UlockLayer(tr, db);

                                        //collect all objectid in the block table record  
                                        IEnumerable<ObjectId> b = btr.Cast<ObjectId>();

                                        #region Collect all Entity
                                        // collect all acdb face 
                                        var AcDbFace = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbFace".ToUpper()).ToList<ObjectId>();

                                        // collect all Acdb Blockreference
                                        var AcdbBlockreference = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbBlockReference".ToUpper()).ToList<ObjectId>();

                                        // collect all AcDb Polyline
                                        var AcDbPolyline = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbPolyline".ToUpper()).ToList<ObjectId>();//

                                        // collect all AcDb Attribute Definition
                                        var AcDbAttributeDefinition = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbAttributeDefinition".ToUpper()).ToList<ObjectId>();

                                        // collect all AcDb _Arc
                                        // var AcDb_Arc = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbArc".ToUpper()).ToList<ObjectId>();//

                                        // collect all AcDb Line
                                        var AcDbLine = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbLine".ToUpper()).ToList<ObjectId>();//

                                        // collect all AcDb Spline
                                        var AcDbSpline = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbSpline".ToUpper()).ToList<ObjectId>();//

                                        // collect all AcDb Hatch
                                        var AcDbHatch = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbHatch".ToUpper()).ToList<ObjectId>();//

                                        // collect all AcDb Arc
                                        var AcDbArc = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbArc".ToUpper()).ToList<ObjectId>();//

                                        // collect all AcDb Leader
                                        var AcDbLeader = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbLeader".ToUpper()).ToList<ObjectId>();//

                                        // collect all AcDb Rotated Dimension 
                                        var AcDbRotatedDimension = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbRotatedDimension".ToUpper()).ToList<ObjectId>();

                                        // collect all AcRx Object
                                        var AcRxObject = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcRxObject".ToUpper()).ToList<ObjectId>();

                                        // collect all AcDb Aligned Dimension
                                        var AcDbAlignedDimension = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbAlignedDimension".ToUpper()).ToList<ObjectId>();

                                        // collect all AcDb MText
                                        var AcDbMText = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbMText".ToUpper()).ToList<ObjectId>();//

                                        //collect all AcDbText
                                        var AcDbText = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbText".ToUpper()).ToList<ObjectId>();//
                                        if (findText(AcDbText, tr, Properties.Settings.Default.SectionKW)) { countSection++; isSection = true; };
                                        if (findText(AcDbText, tr, Properties.Settings.Default.ElevationKW)) { countElevation++; isElevation = true; };
                                        if (findText(AcDbText, tr, Properties.Settings.Default.FloorKW)) { countFloor++; isFloor = true; };
                                        if (isSection)
                                            lstLevels = getlstLevels(AcDbText, tr, Properties.Settings.Default.LevelLayername);
                                        // collect all AcDb Ellipse
                                        var AcDbEllipse = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbEllipse".ToUpper()).ToList<ObjectId>();//

                                        // collect all AcDb 2 Line Angular Dimension
                                        var AcDb2LineAngularDimension = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDb2LineAngularDimension".ToUpper()).ToList<ObjectId>();

                                        // collect all AcDb MLeader
                                        var AcDbMLeader = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbMLeader".ToUpper()).ToList<ObjectId>();//

                                        // collect all AcDb Point
                                        var AcDbPoint = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbPoint".ToUpper()).ToList<ObjectId>();
                                        //collect all circle  AcDbCircle
                                        var AcDbCircle = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbCircle".ToUpper()).ToList<ObjectId>();

                                        // collect all AcDb Solid
                                        var AcDbSolid = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbSolid".ToUpper()).ToList<ObjectId>();

                                        // collect all AcDb Wipeout
                                        var AcDbWipeout = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbWipeout".ToUpper()).ToList<ObjectId>();

                                        // collect all AcDb Mline
                                        var AcDbMline = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbMline".ToUpper()).ToList<ObjectId>();//

                                        #endregion Collect all Entity

                                        #region write Cap file
                                        // write Cap file

                                        try
                                        {

                                            writer.WriteStartElement("AcDbRotatedDimension");
                                            xportAcDbRotatedDimension(AcDbRotatedDimension, writer, tr);//rotated Dimention
                                            writer.WriteEndElement();
                                            writer.WriteStartElement("AcDbLeader");
                                            xportAcDbLeader(AcDbLeader, writer, tr); //Leader
                                            writer.WriteEndElement();
                                            writer.WriteStartElement("AcDbAcDbEllipseRotatedDimension");
                                            xportAcDbEllipse(AcDbEllipse, writer, tr);//Ellipse
                                            writer.WriteEndElement();
                                            writer.WriteStartElement("AcDbAlignedDimension");
                                            xportAcDbAlignedDimension(AcDbAlignedDimension, writer, tr); //Aligned Dimention
                                            writer.WriteEndElement();
                                            writer.WriteStartElement("AcDbHatch");
                                            xportAcDbHatch(AcDbHatch, writer, tr);//hatch
                                            writer.WriteEndElement();
                                            writer.WriteStartElement("AcDbMLeader");
                                            xportAcDbMLeader(AcDbMLeader, writer, tr);//Mleader //pending
                                            writer.WriteEndElement();
                                            writer.WriteStartElement("AcDbArc");
                                            xportAcDbArc(AcDbArc, writer, tr); //arc
                                            writer.WriteEndElement();
                                            writer.WriteStartElement("AcDbMline");
                                            xportAcDbMline(AcDbMline, writer, tr); //mline
                                            writer.WriteEndElement();
                                            writer.WriteStartElement("AcDbCircle");
                                            xportAcDbCircle(AcDbCircle, writer, tr);//Circle
                                            writer.WriteEndElement();
                                            writer.WriteStartElement("AcDbMText");
                                            xportMtext(AcDbMText, writer, tr); //Mtext
                                            writer.WriteEndElement();
                                            writer.WriteStartElement("AcDbText");
                                            xportdbText(AcDbText, writer, tr);//Text
                                            writer.WriteEndElement();
                                            writer.WriteStartElement("AcDbLine");
                                            Xportline(AcDbLine, writer, tr);//line
                                            writer.WriteEndElement();
                                            writer.WriteStartElement("AcDbPolyline");
                                            xportPolyline(AcDbPolyline, writer, tr);//polyline
                                            writer.WriteEndElement();
                                            writer.WriteStartElement("AcDbSpline");
                                            xportSpline(AcDbSpline, writer, tr);//spline
                                            writer.WriteEndElement();
                                            //XportWall(AcDbLine, writer, tr);

                                        }
                                        catch (Exception ex)
                                        { Microsoft.WindowsAPICodePack.Dialogs.TaskDialog.Show(ex.Message); }
                                        #endregion write Cap file

                                        tr.Commit();

                                    }
                                }
                            writer.WriteEndElement();
                        }
                        Mouse.OverrideCursor = System.Windows.Input.Cursors.Arrow;
                    }

                    writer.WriteEndElement();

                    MessageBox.Show(@"Total " + filecount + " *.dwg files are deducted! in this we have below : \n\n"
                        + "Section " + countSection + "*.dwg files.\n"
                        + "Elevation " + countElevation + "*.dwg files.\n"
                        + "Plan " + countFloor + "*.dwg files.\n", "CapInfo");
                }
                //Microsoft.WindowsAPICodePack.Dialogs.TaskDialog.Show("CapInfo", "Completed");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "CapInfo");

            }
            finally
            {
                //btnScan.IsEnabled = true;
                Mouse.OverrideCursor = System.Windows.Input.Cursors.Arrow;
            }
        }
        private static bool IsFileWritable(string filename)
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
        static BitmapImage getBitmap(string imageName, int Height, int Width)
        {
            BitmapImage image = new BitmapImage();
            image.BeginInit();
            image.StreamSource = Assembly.GetExecutingAssembly().GetManifestResourceStream(imageName);
            image.UriSource = new Uri(imageName);
            image.DecodePixelHeight = Height;
            image.DecodePixelWidth = Width;
            image.EndInit();
            return image;
        }
        public static event EventHandler PaletteSetClosed;
        public class RibbonButton : System.Windows.Input.ICommand
        {
            private string[] lstFiles;
            private int filecount;

            public bool CanExecute(object parameter)
            {
                return true;
            }
            public event EventHandler CanExecuteChanged;

            public void Execute(object parameter)
            {
                Autodesk.Windows.RibbonButton button = parameter as Autodesk.Windows.RibbonButton;

                if (button.Id == "btnCap")
                {
                    string filename = string.Empty;


                    var doc = Application.DocumentManager.MdiActiveDocument;
                    // var db = doc.Database;
                    var ed = doc.Editor;
                    HostApplicationServices hs = HostApplicationServices.Current;

                    List<string> _fileslist = new List<string>();
                    using (var fbd = new FolderBrowserDialog())
                    {
                        DialogResult result = fbd.ShowDialog();
                        if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                        {
                            lstFiles = Directory.GetFiles(fbd.SelectedPath, "*.dwg*");
                        }
                    }
                    filecount = lstFiles.Count();
                    Microsoft.WindowsAPICodePack.Dialogs.TaskDialog.Show("Total " + filecount + " *.dwg files are deducted!.");
                    string currentDirectory = Path.GetDirectoryName(@"C:\");
                    SaveFileDialog openFileDialog1 = new SaveFileDialog
                    {
                        InitialDirectory = currentDirectory,
                        Title = "Browse Cap Files",
                        CheckFileExists = false,
                        CheckPathExists = true,
                        DefaultExt = "cap",
                        Filter = "Cap files (*.cap)|*.cap",
                        FilterIndex = 2,
                        RestoreDirectory = true,

                    };

                    if (openFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        filename = openFileDialog1.FileName;
                        //Microsoft.WindowsAPICodePack.Dialogs.TaskDialog.Show("Total " + filecount1 + " *.dwg files are deducted!.");


                        foreach (string _file in lstFiles)
                        {
                            //Cursor.Current = Cursors.WaitCursor;

                            if (_file != null && _file != "")
                            {
                                //string currentDirectory = Path.GetDirectoryName(@"C:\");
                                using (var acDOCLOCK = doc.LockDocument())
                                {

                                    var dwgfilename = Path.GetFileNameWithoutExtension(lstFiles[0]).ToString();
                                    var fisrtfilename = Path.GetFileNameWithoutExtension(_file).ToString();

                                    using (Database db = new Database(false, true))
                                    {
                                        db.ReadDwgFile(_file, System.IO.FileShare.ReadWrite, false, "");
                                        using (var tr = db.TransactionManager.StartOpenCloseTransaction())
                                        {
                                            //reading all details from the model space
                                            var modelSpace = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                                            var brclass = RXObject.GetClass(typeof(BlockReference));
                                            //collecting all blocks in the Model 
                                            var blocks = modelSpace.Cast<ObjectId>().Where(id => id.ObjectClass == brclass).Select(id => (BlockReference)tr.GetObject(id, OpenMode.ForRead))
                                                .GroupBy(br => ((BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead)).Name);

                                            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);


                                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                                            // UlockLayer(tr, db);

                                            //collect all objectid in the block table record  
                                            IEnumerable<ObjectId> b = btr.Cast<ObjectId>();
                                            var AcdbBlockreference = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbBlockReference".ToUpper()).ToList<ObjectId>();
                                            //   ExplodeBlock(tr, db, AcdbBlockreference);
                                            //collect all block names in list
                                            List<ObjectId> collection = new List<ObjectId>();
                                            collection.Clear();

                                            List<string> _Entityname = new List<string>();
                                            List<string> collectionname = new List<string>();

                                            #region Collect all Entity
                                            // collect all acdb face 
                                            var AcDbFace = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbFace".ToUpper()).ToList<ObjectId>();

                                            // collect all Acdb Blockreference

                                            // collect all AcDb Polyline
                                            var AcDbPolyline = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbPolyline".ToUpper()).ToList<ObjectId>();//

                                            // collect all AcDb Attribute Definition
                                            var AcDbAttributeDefinition = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbAttributeDefinition".ToUpper()).ToList<ObjectId>();

                                            // collect all AcDb _Arc
                                            // var AcDb_Arc = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbArc".ToUpper()).ToList<ObjectId>();//

                                            // collect all AcDb Line
                                            var AcDbLine = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbLine".ToUpper()).ToList<ObjectId>();//

                                            // collect all AcDb Spline
                                            var AcDbSpline = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbSpline".ToUpper()).ToList<ObjectId>();//

                                            // collect all AcDb Hatch
                                            var AcDbHatch = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbHatch".ToUpper()).ToList<ObjectId>();//

                                            // collect all AcDb Arc
                                            var AcDbArc = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbArc".ToUpper()).ToList<ObjectId>();//

                                            // collect all AcDb Leader
                                            var AcDbLeader = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbLeader".ToUpper()).ToList<ObjectId>();//

                                            // collect all AcDb Rotated Dimension 
                                            var AcDbRotatedDimension = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbRotatedDimension".ToUpper()).ToList<ObjectId>();

                                            // collect all AcRx Object
                                            var AcRxObject = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcRxObject".ToUpper()).ToList<ObjectId>();

                                            // collect all AcDb Aligned Dimension
                                            var AcDbAlignedDimension = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbAlignedDimension".ToUpper()).ToList<ObjectId>();

                                            // collect all AcDb MText
                                            var AcDbMText = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbMText".ToUpper()).ToList<ObjectId>();//

                                            //collect all AcDbText
                                            var AcDbText = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbText".ToUpper()).ToList<ObjectId>();//
                                            Microsoft.WindowsAPICodePack.Dialogs.TaskDialog.Show("CapInfo", findText(AcDbText, tr, "Verdieping").ToString());
                                            // collect all AcDb Ellipse
                                            var AcDbEllipse = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbEllipse".ToUpper()).ToList<ObjectId>();//

                                            // collect all AcDb 2 Line Angular Dimension
                                            var AcDb2LineAngularDimension = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDb2LineAngularDimension".ToUpper()).ToList<ObjectId>();

                                            // collect all AcDb MLeader
                                            var AcDbMLeader = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbMLeader".ToUpper()).ToList<ObjectId>();//

                                            // collect all AcDb Point
                                            var AcDbPoint = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbPoint".ToUpper()).ToList<ObjectId>();
                                            //collect all circle  AcDbCircle
                                            var AcDbCircle = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbCircle".ToUpper()).ToList<ObjectId>();

                                            // collect all AcDb Solid
                                            var AcDbSolid = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbSolid".ToUpper()).ToList<ObjectId>();

                                            // collect all AcDb Wipeout
                                            var AcDbWipeout = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbWipeout".ToUpper()).ToList<ObjectId>();

                                            // collect all AcDb Mline
                                            var AcDbMline = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbMline".ToUpper()).ToList<ObjectId>();//

                                            #endregion Collect all Entity

                                            var collcount = collection.Count;
                                            var entity = b.Count();
                                            // Create an XmlWriterSettings object with the correct options.
                                            System.Xml.XmlWriterSettings settings = new System.Xml.XmlWriterSettings();
                                            settings.Indent = true;
                                            settings.IndentChars = "    ";
                                            settings.OmitXmlDeclaration = false;
                                            settings.Encoding = System.Text.Encoding.UTF8;


                                            try
                                            {
                                                #region cmdd
                                                //using (System.Xml.XmlWriter writer = System.Xml.XmlWriter.Create(filename, settings))
                                                //{
                                                //    writer.WriteStartDocument();
                                                //    writer.WriteStartElement("Export");
                                                //    writer.WriteStartElement("CAD");

                                                //    for (int i = 0; i < filecount1; i++)
                                                //    {
                                                //        xportAcDbRotatedDimension(AcDbRotatedDimension, writer, tr);//rotated Dimention
                                                //        xportAcDbLeader(AcDbLeader, writer, tr); //Leader
                                                //        xportAcDbEllipse(AcDbEllipse, writer, tr);//Ellipse
                                                //        xportAcDbAlignedDimension(AcDbAlignedDimension, writer, tr); //Aligned Dimention
                                                //        xportAcDbHatch(AcDbHatch, writer, tr);//hatch
                                                //        xportAcDbMLeader(AcDbMLeader, writer, tr);//Mleader //pending
                                                //        xportAcDbArc(AcDbArc, writer, tr); //arc
                                                //        xportAcDbMline(AcDbMline, writer, tr); //mline
                                                //        xportAcDbCircle(AcDbCircle, writer, tr);//Circle
                                                //        xportMtext(AcDbMText, writer, tr); //Mtext
                                                //        xportdbText(AcDbText, writer, tr);//Text
                                                //        Xportline(AcDbLine, writer, tr);//line
                                                //        xportPolyline(AcDbPolyline, writer, tr);//polyline
                                                //        xportSpline(AcDbSpline, writer, tr);//spline                 
                                                //        //XportWall(AcDbLine, writer, tr);
                                                //    }
                                                //    writer.WriteEndElement();
                                                //    writer.WriteStartElement("CAP");
                                                //    Wallhatchdetails(AcDbHatch, writer, tr);

                                                //    writer.Flush();
                                                //    writer.Close();
                                                //}
                                                #endregion cmdd
                                                tr.Commit();

                                            }
                                            catch (Exception ex) { MessageBox.Show("", "" + ex.Message); }
                                        }
                                    }
                                }

                            }
                        }
                        #region cmd

                        //string cap = Path.Combine(Path.GetFullPath(System.Reflection.Assembly.GetExecutingAssembly().Location)
                        //  .Remove(Path.GetFullPath(System.Reflection.Assembly.GetExecutingAssembly().Location).Count()
                        //  - Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location).Count()), @"CAPFILE\2D-3D.cap");
                        //if (System.IO.File.Exists(cap))
                        //    System.IO.File.Copy(cap, filename);
                        #endregion cmd
                        //Microsoft.WindowsAPICodePack.Dialogs.TaskDialog.Show("Cap Exported " + filecount1 + " dwg Files", "Completed");
                        Microsoft.WindowsAPICodePack.Dialogs.TaskDialog.Show("CapInfo", "Completed");
                    }
                }
                else if (button.Id == "btnEcap")
                {
                    string filename = string.Empty;
                    countHatchs = 0;
                    lstLevelAreas = new List<string>();
                    lstLevels = new List<double>();
                    levelHatchs = new SortedList<string, List<Curve>>();
                    wallPolyline = new SortedList<string, List<Curve>>();


                    var doc = Application.DocumentManager.MdiActiveDocument;
                    // var db = doc.Database;
                    var ed = doc.Editor;
                    HostApplicationServices hs = HostApplicationServices.Current;

                    List<string> _fileslist = new List<string>();
                    using (var fbd = new FolderBrowserDialog())
                    {
                        fbd.SelectedPath = ECAP.Properties.Settings.Default.catchEcapPath;
                        DialogResult result = fbd.ShowDialog();
                        if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                        {
                            ECAP.Properties.Settings.Default.catchEcapPath = fbd.SelectedPath;
                            ECAP.Properties.Settings.Default.Save();
                            lstFiles = Directory.GetFiles(fbd.SelectedPath, "*.dwg*");
                        }
                    }
                    filecount = lstFiles.Count();
                    int countSection = 0;
                    int countElevation = 0;
                    int countFloor = 0;



                    SaveFileDialog openFileDialog1 = new SaveFileDialog
                    {
                        InitialDirectory = ECAP.Properties.Settings.Default.catchEcapPath,
                        Title = "Browse Cap Files",
                        CheckFileExists = false,
                        CheckPathExists = true,
                        DefaultExt = "cap",
                        Filter = "Cap files (*.cap)|*.cap",
                        FilterIndex = 2,
                        RestoreDirectory = true,

                    };

                    if (openFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        filename = openFileDialog1.FileName;

                        try
                        {
                            System.Xml.XmlWriterSettings settings = new System.Xml.XmlWriterSettings();
                            settings.Indent = true;
                            settings.IndentChars = "    ";
                            settings.OmitXmlDeclaration = false;
                            settings.Encoding = System.Text.Encoding.UTF8;


                            using (System.Xml.XmlWriter writer = System.Xml.XmlWriter.Create(filename, settings))
                            {
                                writer.WriteStartDocument();
                                writer.WriteStartElement("Export");
                                writer.WriteStartElement("CAD");

                                foreach (string _file in lstFiles)
                                {
                                    isSection = false;
                                    isElevation = false;
                                    isFloor = false;
                                    //Cursor.Current = Cursors.WaitCursor;
                                    currentFilename = System.IO.Path.GetFileNameWithoutExtension(_file).ToString();
                                    if (_file != null && _file != "")
                                    {
                                        writer.WriteStartElement("File");
                                        writer.WriteAttributeString("FileName", System.IO.Path.GetFileNameWithoutExtension(_file).ToString());
                                        writer.WriteAttributeString("FilePath", _file.ToString());
                                        using (var acDOCLOCK = doc.LockDocument())
                                        using (Database db = new Database(false, true))
                                        {
                                            db.ReadDwgFile(_file, System.IO.FileShare.ReadWrite, false, "");
                                            using (var tr = db.TransactionManager.StartOpenCloseTransaction())
                                            {
                                                //reading all details from the model space
                                                var modelSpace = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                                                var brclass = RXObject.GetClass(typeof(BlockReference));
                                                //collecting all blocks in the Model 
                                                var blocks = modelSpace.Cast<ObjectId>().Where(id => id.ObjectClass == brclass).Select(id => (BlockReference)tr.GetObject(id, OpenMode.ForRead))
                                                    .GroupBy(br => ((BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead)).Name);

                                                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);


                                                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                                                // UlockLayer(tr, db);

                                                //collect all objectid in the block table record  
                                                IEnumerable<ObjectId> b = btr.Cast<ObjectId>();

                                                #region Collect all Entity
                                                // collect all acdb face 
                                                var AcDbFace = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbFace".ToUpper()).ToList<ObjectId>();

                                                // collect all Acdb Blockreference
                                                var AcdbBlockreference = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbBlockReference".ToUpper()).ToList<ObjectId>();

                                                // collect all AcDb Polyline
                                                var AcDbPolyline = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbPolyline".ToUpper()).ToList<ObjectId>();//

                                                // collect all AcDb Attribute Definition
                                                var AcDbAttributeDefinition = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbAttributeDefinition".ToUpper()).ToList<ObjectId>();

                                                // collect all AcDb _Arc
                                                // var AcDb_Arc = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbArc".ToUpper()).ToList<ObjectId>();//

                                                // collect all AcDb Line
                                                var AcDbLine = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbLine".ToUpper()).ToList<ObjectId>();//

                                                // collect all AcDb Spline
                                                var AcDbSpline = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbSpline".ToUpper()).ToList<ObjectId>();//

                                                // collect all AcDb Hatch
                                                var AcDbHatch = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbHatch".ToUpper()).ToList<ObjectId>();//

                                                // collect all AcDb Arc
                                                var AcDbArc = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbArc".ToUpper()).ToList<ObjectId>();//

                                                // collect all AcDb Leader
                                                var AcDbLeader = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbLeader".ToUpper()).ToList<ObjectId>();//

                                                // collect all AcDb Rotated Dimension 
                                                var AcDbRotatedDimension = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbRotatedDimension".ToUpper()).ToList<ObjectId>();

                                                // collect all AcRx Object
                                                var AcRxObject = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcRxObject".ToUpper()).ToList<ObjectId>();

                                                // collect all AcDb Aligned Dimension
                                                var AcDbAlignedDimension = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbAlignedDimension".ToUpper()).ToList<ObjectId>();

                                                // collect all AcDb MText
                                                var AcDbMText = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbMText".ToUpper()).ToList<ObjectId>();//

                                                //collect all AcDbText
                                                var AcDbText = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbText".ToUpper()).ToList<ObjectId>();//
                                                if (findText(AcDbText, tr, ECAP.Properties.Settings.Default.SectionKW)) { countSection++; isSection = true; };
                                                if (findText(AcDbText, tr, ECAP.Properties.Settings.Default.ElevationKW)) { countElevation++; isElevation = true; };
                                                if (findText(AcDbText, tr, ECAP.Properties.Settings.Default.FloorKW)) { countFloor++; isFloor = true; };
                                                if (isSection)
                                                    lstLevels = getlstLevels(AcDbText, tr, ECAP.Properties.Settings.Default.LevelLayername);
                                                // collect all AcDb Ellipse
                                                var AcDbEllipse = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbEllipse".ToUpper()).ToList<ObjectId>();//

                                                // collect all AcDb 2 Line Angular Dimension
                                                var AcDb2LineAngularDimension = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDb2LineAngularDimension".ToUpper()).ToList<ObjectId>();

                                                // collect all AcDb MLeader
                                                var AcDbMLeader = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbMLeader".ToUpper()).ToList<ObjectId>();//

                                                // collect all AcDb Point
                                                var AcDbPoint = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbPoint".ToUpper()).ToList<ObjectId>();
                                                //collect all circle  AcDbCircle
                                                var AcDbCircle = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbCircle".ToUpper()).ToList<ObjectId>();

                                                // collect all AcDb Solid
                                                var AcDbSolid = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbSolid".ToUpper()).ToList<ObjectId>();

                                                // collect all AcDb Wipeout
                                                var AcDbWipeout = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbWipeout".ToUpper()).ToList<ObjectId>();

                                                // collect all AcDb Mline
                                                var AcDbMline = b.Where(id => id.ObjectClass.Name.ToUpper() == "AcDbMline".ToUpper()).ToList<ObjectId>();//

                                                #endregion Collect all Entity

                                                #region write Cap file
                                                // write Cap file

                                                try
                                                {

                                                    writer.WriteStartElement("AcDbRotatedDimension");
                                                    xportAcDbRotatedDimension(AcDbRotatedDimension, writer, tr);//rotated Dimention
                                                    writer.WriteEndElement();
                                                    writer.WriteStartElement("AcDbLeader");
                                                    xportAcDbLeader(AcDbLeader, writer, tr); //Leader
                                                    writer.WriteEndElement();
                                                    writer.WriteStartElement("AcDbAcDbEllipseRotatedDimension");
                                                    xportAcDbEllipse(AcDbEllipse, writer, tr);//Ellipse
                                                    writer.WriteEndElement();
                                                    writer.WriteStartElement("AcDbAlignedDimension");
                                                    xportAcDbAlignedDimension(AcDbAlignedDimension, writer, tr); //Aligned Dimention
                                                    writer.WriteEndElement();
                                                    writer.WriteStartElement("AcDbHatch");
                                                    xportAcDbHatch(AcDbHatch, writer, tr);//hatch
                                                    writer.WriteEndElement();
                                                    writer.WriteStartElement("AcDbMLeader");
                                                    xportAcDbMLeader(AcDbMLeader, writer, tr);//Mleader //pending
                                                    writer.WriteEndElement();
                                                    writer.WriteStartElement("AcDbArc");
                                                    xportAcDbArc(AcDbArc, writer, tr); //arc
                                                    writer.WriteEndElement();
                                                    writer.WriteStartElement("AcDbMline");
                                                    xportAcDbMline(AcDbMline, writer, tr); //mline
                                                    writer.WriteEndElement();
                                                    writer.WriteStartElement("AcDbCircle");
                                                    xportAcDbCircle(AcDbCircle, writer, tr);//Circle
                                                    writer.WriteEndElement();
                                                    writer.WriteStartElement("AcDbMText");
                                                    xportMtext(AcDbMText, writer, tr); //Mtext
                                                    writer.WriteEndElement();
                                                    writer.WriteStartElement("AcDbText");
                                                    xportdbText(AcDbText, writer, tr);//Text
                                                    writer.WriteEndElement();
                                                    writer.WriteStartElement("AcDbLine");
                                                    Xportline(AcDbLine, writer, tr);//line
                                                    writer.WriteEndElement();
                                                    writer.WriteStartElement("AcDbPolyline");
                                                    xportPolyline(AcDbPolyline, writer, tr);//polyline
                                                    writer.WriteEndElement();
                                                    writer.WriteStartElement("AcDbSpline");
                                                    xportSpline(AcDbSpline, writer, tr);//spline
                                                    writer.WriteEndElement();
                                                    //XportWall(AcDbLine, writer, tr);

                                                }
                                                catch (Exception ex)
                                                { Microsoft.WindowsAPICodePack.Dialogs.TaskDialog.Show(ex.Message); }
                                                #endregion write Cap file

                                                tr.Commit();

                                            }
                                        }
                                        writer.WriteEndElement();
                                    }
                                    //Cursor.Current = Cursors.Default;
                                }

                                writer.WriteEndElement();
                                writer.WriteStartElement("CAP");
                                //Wallhatchdetails(AcDbHatch, writer, tr);
                                writer.WriteStartElement("Levels");
                                SortedList<string, string> listleveldetails = new SortedList<string, string>();
                                int countLevel = 0;
                                frmLevel frmLevel = new frmLevel(lstLevels, levelHatchs);
                                frmLevel.ShowDialog();
                                if (lstLevels.Count > 0)
                                {
                                    foreach (double levelVal in lstLevels)
                                    {
                                        if (levelVal >= 0)
                                        {

                                            writer.WriteStartElement("Level" + countLevel);
                                            string levelname = countLevel == 0 ? "Ground Floor" : (countLevel == 1 ? "1st Floor" : (countLevel == 2 ? "2nd Floor" : "3rd Floor"));
                                            writer.WriteElementString("Level_Name", levelname);
                                            writer.WriteElementString("Level_Number", countLevel.ToString());
                                            writer.WriteElementString("Level_Value", levelVal.ToString());
                                            writer.WriteEndElement();
                                            if (!listleveldetails.ContainsKey(levelname.ToLower().Replace(" ", "")))
                                                listleveldetails.Add(levelname.ToLower().Replace(" ", ""), levelVal.ToString());
                                            countLevel++;
                                        }
                                    }
                                }
                                writer.WriteEndElement();

                                writer.WriteStartElement("Floors");

                                foreach (string levelkey in levelHatchs.Keys)
                                {
                                    List<Curve> curves = levelHatchs[levelkey];
                                    writer.WriteStartElement(levelkey.Split(';')[0].ToString());
                                    string lvname = levelkey.Split(';')[1].ToString();
                                    writer.WriteElementString("Level_Name", lvname);

                                    if (listleveldetails.ContainsKey(lvname))
                                    {
                                        writer.WriteElementString("Level_Number", (lvname == "groundfloor" ? "0" : (lvname == "1stfloor" ? "1" : (lvname == "2ndfloor" ? "2" : "3"))));
                                        writer.WriteElementString("Level_Value", listleveldetails[lvname]);
                                    }
                                    int xcount = 0;
                                    double zValue = 0;
                                    if (listleveldetails.ContainsKey(lvname))
                                        double.TryParse(listleveldetails[lvname], out zValue);
                                    foreach (var curve in curves)
                                    {
                                        writer.WriteElementString("cPoint", "" + new Point3d(curve.StartPoint.X, curve.StartPoint.Y, zValue).ToString());
                                        xcount++;
                                        if (xcount == curves.Count)
                                            writer.WriteElementString("cPoint", "" + new Point3d(curve.EndPoint.X, curve.EndPoint.Y, zValue).ToString());
                                    }
                                    writer.WriteEndElement();
                                }
                                writer.WriteEndElement();
                                writer.WriteStartElement("Walls");

                                foreach (string wallkey in wallPolyline.Keys)
                                {
                                    List<Curve> curves = wallPolyline[wallkey];
                                    writer.WriteStartElement(wallkey.Split(';')[0].ToString());
                                    string lvname = wallkey.Split(';')[1].ToString();
                                    writer.WriteElementString("Level_Name", lvname);

                                    if (listleveldetails.ContainsKey(lvname))
                                    {
                                        writer.WriteElementString("Level_Number", (lvname == "groundfloor" ? "0" : (lvname == "1stfloor" ? "1" : (lvname == "2ndfloor" ? "2" : "3"))));
                                        writer.WriteElementString("Level_Value", listleveldetails[lvname]);
                                    }
                                    int xcount = 0;
                                    double zValue = 0;
                                    if (listleveldetails.ContainsKey(lvname))
                                        double.TryParse(listleveldetails[lvname], out zValue);
                                    foreach (var curve in curves)
                                    {
                                        writer.WriteElementString("cPoint", "" + new Point3d(curve.StartPoint.X, curve.StartPoint.Y, zValue).ToString());
                                        xcount++;
                                        if (xcount == curves.Count)
                                            writer.WriteElementString("cPoint", "" + new Point3d(curve.EndPoint.X, curve.EndPoint.Y, zValue).ToString());
                                    }
                                    writer.WriteEndElement();
                                }
                                writer.WriteEndElement();

                                writer.Flush();
                                writer.Close();
                            }

                            Microsoft.WindowsAPICodePack.Dialogs.TaskDialog.Show("CapInfo", @"Total " + filecount + " *.dwg files are deducted! in this we have below : \n\n"
                                + "Section " + countSection + "*.dwg files.\n"
                                + "Elevation " + countElevation + "*.dwg files.\n"
                                + "Plan " + countFloor + "*.dwg files.\n");
                            //Microsoft.WindowsAPICodePack.Dialogs.TaskDialog.Show("CapInfo", "Completed");
                        }
                        catch (Exception ex) { Microsoft.WindowsAPICodePack.Dialogs.TaskDialog.Show("CapInfo", ex.Message); }
                    }
                }
            }


            private static void Wallhatchdetails(List<ObjectId> acDbHatch, XmlWriter writer, OpenCloseTransaction tr)
            {

                foreach (var objid in acDbHatch)
                {

                    Entity enttxt = (Entity)tr.GetObject(objid, OpenMode.ForWrite);

                    if (enttxt.Layer.ToUpper().Contains("Wall".ToUpper()))
                    {
                        Hatch hatch = tr.GetObject(objid, OpenMode.ForRead) as Hatch;
                        int nLoops = hatch.NumberOfLoops;

                        for (int i = 0; i < nLoops; i++)
                        {
                            HatchLoopTypes hlt = hatch.LoopTypeAt(i);
                            HatchLoop hatchLoop = hatch.GetLoopAt(i);
                            Curve2dCollection curves = hatchLoop.Curves;
                            if (curves != null)
                            {
                                List<double> maximum = new List<double>();
                                //collect length in hatch
                                foreach (Curve2d curve in curves)
                                {
                                    var length = ((LineSegment2d)curve).Length;
                                    maximum.Add(length);
                                }
                                //get the maximum number in list
                                double first = maximum.Max();
                                var maxlength = maximum.Count;
                                //get the second maxium number
                                var second = maximum.OrderByDescending(x => x).Distinct().Skip(1).First();
                                double third = 0;
                                if (maxlength == 5)
                                {
                                    third = maximum.OrderByDescending(x => x).Distinct().Skip(2).First();
                                }
                                foreach (Curve2d curve in curves)
                                {
                                    var length = ((LineSegment2d)curve).Length;

                                    if ((length == first) || (length == second) || (length == third))
                                    {
                                        writer.WriteStartElement("Hatch");
                                        var startpoint = curve.StartPoint;
                                        var Endpoint = curve.EndPoint;
                                        var MidPoint = ((Autodesk.AutoCAD.Geometry.LineSegment2d)curve).MidPoint;
                                        writer.WriteElementString("StartPoint", "" + startpoint.ToString());
                                        writer.WriteElementString("EndPoint", "" + Endpoint.ToString());
                                        writer.WriteElementString("Height", "" + "2500".ToString());
                                        writer.WriteEndElement();
                                    }

                                }

                            }

                        }
                    }

                }


            }
        }

        private static void XportWall(IEnumerable<ObjectId> Wall, XmlWriter writer, OpenCloseTransaction tr)
        {
            try
            {
                foreach (var EntName in Wall)
                {
                    string content = string.Empty;
                    Entity enttxt = (Entity)tr.GetObject(EntName, OpenMode.ForWrite);
                    int ct = Wall.Count();
                    if (enttxt.Layer.ToUpper().Contains("WALL".ToUpper()))
                    {
                        if (enttxt is Line)
                        {
                            Line line = ((Line)enttxt);
                            var lengh = line.Length;
                            layerNAME = line.Layer;
                            var startpoint = line.StartPoint;
                            var endpoint = line.EndPoint;
                            var elename = EntName.ObjectClass.Name;
                            var lineweight = line.LineWeight;
                            var lineType = line.Angle;
                            var meterial = line.Material;
                            var height = 2500;
                            writer.WriteStartElement("Wall");
                            writer.WriteElementString("ElementID", "" + EntName.ToString());
                            writer.WriteElementString("ElementName", "" + elename.ToString());
                            writer.WriteElementString("LayerName", "" + layerNAME.ToString());
                            writer.WriteElementString("Length", "" + lengh.ToString());
                            writer.WriteElementString("LineWeight", "" + lineweight.ToString());
                            writer.WriteElementString("LineType", "" + lineType.ToString());
                            writer.WriteElementString("Material", "" + meterial.ToString());
                            writer.WriteElementString("StartPoint", "" + startpoint.ToString());
                            writer.WriteElementString("EndPoint", "" + endpoint.ToString());
                            writer.WriteElementString("Height", "" + height.ToString());
                            writer.WriteEndElement();
                        }
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("", "" + ex.Message); }
        }

        //private static void XportWallhatch(IEnumerable<ObjectId> Wall, XmlWriter writer, OpenCloseTransaction tr)
        //{
        //    try
        //    {
        //        foreach (var EntName in Wall)
        //        {
        //            string content = string.Empty;
        //            Entity enttxt = (Entity)tr.GetObject(EntName, OpenMode.ForWrite);
        //            int ct = Wall.Count();
        //            if (enttxt.Layer.ToUpper().Contains("WALL".ToUpper()))
        //            {
        //                if (enttxt is Line)
        //                {
        //                    Hatch htch = ((Hatch)enttxt);
        //                    //var lengh = Hatch.Length;
        //                    layerNAME = htch.Layer;
        //                    var startpoint = htch.StartPoint;
        //                    var endpoint = htch.EndPoint;
        //                    var elename = EntName.ObjectClass.Name;
        //                    var lineweight = htch.LineWeight;
        //                    //var lineType = Hatch.Angle;
        //                    var meterial = htch.Material;
        //                    var height = 2500;
        //                    writer.WriteStartElement("Wall");
        //                    writer.WriteElementString("ElementID", "" + EntName.ToString());
        //                    writer.WriteElementString("ElementName", "" + elename.ToString());
        //                    writer.WriteElementString("LayerName", "" + layerNAME.ToString());
        //                   // writer.WriteElementString("Length", "" + lengh.ToString());
        //                    writer.WriteElementString("LineWeight", "" + lineweight.ToString());
        //                   // writer.WriteElementString("LineType", "" + lineType.ToString());
        //                    writer.WriteElementString("Material", "" + meterial.ToString());
        //                    writer.WriteElementString("StartPoint", "" + startpoint.ToString());
        //                    writer.WriteElementString("EndPoint", "" + endpoint.ToString());
        //                    writer.WriteElementString("Height", "" + height.ToString());
        //                    writer.WriteEndElement();
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex) { MessageBox.Show("", "" + ex.Message); }
        //}
        private static void DecryptCapFile(string filename)
        {
            string text = File.ReadAllText(filename);
            byte[] b;
            string decrypted;
            try
            {
                b = Convert.FromBase64String(text);
                decrypted = System.Text.ASCIIEncoding.ASCII.GetString(b);
                using (StreamWriter writetext = new StreamWriter(filename))
                {
                    writetext.WriteLine(decrypted);
                }
            }
            catch (FormatException fe)
            {
                decrypted = "";
            }

        }
        private static void EncryptCapFile(string filename)
        {

            string text = File.ReadAllText(filename);
            byte[] b = System.Text.ASCIIEncoding.ASCII.GetBytes(text);
            string encrypted = Convert.ToBase64String(b);
            using (StreamWriter writetext = new StreamWriter(filename))
            {
                writetext.WriteLine(encrypted);
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
                    }
                }
                using (StreamWriter writetext = new StreamWriter(filename))
                {
                    writetext.WriteLine(text);
                }
            }
            catch (Exception ex) { MessageBox.Show("ExportCap", "" + ex.Message + ex.StackTrace); }
        }
        private static void ExplodeBlock(OpenCloseTransaction tr, Database db, List<ObjectId> acdbBlockreference)
        {
            // for the explode operation, as it's non-destructive
            foreach (ObjectId id in acdbBlockreference)
            {
                DBObjectCollection objs = new DBObjectCollection();
                //handle as entity and explode
                Entity ent = (Entity)tr.GetObject(id, OpenMode.ForRead);
                ent.Explode(objs);
                //erase Block
                ent.UpgradeOpen();
                ent.Erase();
            }
        }
        //export Aligned Dimention
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

    }

}
