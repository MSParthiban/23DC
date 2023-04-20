using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Media;
using Color = System.Drawing.Color;
using Pen = System.Drawing.Pen;

namespace ECAP
{
    public partial class frmLevel : Form
    {      
        public static Pen redPen = new Pen(Color.FromArgb(255, 0, 0));
        public static Pen greenPen = new Pen(Color.FromArgb(0, 255, 0));
        public static Pen bluePen = new Pen(Color.FromArgb( 0, 0, 255));
        private static SortedList<string, List<Curve>> _levelHatchs = new SortedList<string, List<Curve>>();
        public frmLevel()
        {
            InitializeComponent();
        }
        public frmLevel(List<double> lstLevels, SortedList<string, List<Curve>> levelHatchs)
        { 
            _levelHatchs = levelHatchs;
            InitializeComponent();
          
            int countLevel = 0;
            foreach (double levelVal in lstLevels)
            {
                if (levelVal >= 0)
                {
                    //DataGridRow dataGridRow = new DataGridRow();
                    //dataGridRow.[0] = countLevel;
                    //dataGridRow[1] = levelVal.ToString();
                    dataGridViewLEVEL.Rows.Add("Level "+ countLevel, levelVal.ToString());
                    countLevel++;
                }
            }
      
           
        }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            ECAP.EcapLogin.lstLevels = new List<double>();
            foreach (DataGridViewRow row in dataGridViewLEVEL.Rows)
            {
                if (row.Cells[1].Value == null) continue;
                double levelVal = 0;
                double.TryParse(row.Cells[1].Value.ToString(), out levelVal);
                ECAP.EcapLogin.lstLevels.Add(levelVal);
            }
            //this.Close();
           
        }

        private void picCanvas_Paint(object sender, PaintEventArgs e)
        {
            foreach (string levelkey in _levelHatchs.Keys)
            {
                List<Curve> curves = _levelHatchs[levelkey];
                List<PointF> lstpts = new List<PointF>();
                int xcount = 0;
                foreach (var curve in curves)
                {
                    lstpts.Add(new PointF((float)curve.StartPoint.X, (float)curve.StartPoint.Y));
                  
                   // writer.WriteElementString("cPoint", "" + new Point3d(curve.StartPoint.X, curve.StartPoint.Y, zValue).ToString());
                    xcount++;
                    if (xcount == curves.Count)
                        lstpts.Add(new PointF((float)curve.EndPoint.X, (float)curve.EndPoint.Y));
                }
                e.Graphics.DrawPolygon(redPen, lstpts.ToArray());
            }

        }
    }
}
