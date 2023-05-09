using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace _23DC
{
    /// <summary>
    /// Interaction logic for AutoResizingCanvas.xaml
    /// </summary>

    public partial class AutoResizingCanvas : UserControl
    {
        static double currentSize_X;

        static double currentSize_Y;

        Canvas panel;



        PointCollectionViewModel pcvm;

        public AutoResizingCanvas()

        {

            InitializeComponent();

        }





        private void cvMap_SizeChanged(object sender, SizeChangedEventArgs e)

        {

            panel = sender as Canvas;

            currentSize_X = e.NewSize.Width;

            currentSize_Y = e.NewSize.Height;

            pcvm = new PointCollectionViewModel(currentSize_X, currentSize_Y);

            base.DataContext = pcvm;

        }



        private void cvMap_Loaded(object sender, RoutedEventArgs e)

        {

            panel = sender as Canvas;

            currentSize_X = panel.ActualWidth;

            currentSize_Y = panel.ActualHeight;

            pcvm = new PointCollectionViewModel(currentSize_X, currentSize_Y);

            base.DataContext = pcvm;

        }

    }

    public class PointCollectionViewModel

    {

        private List<PointViewModel> viewModels;

        public PointCollectionViewModel(double currentSize_X, double currentSize_Y)

        {

            this.viewModels = new List<PointViewModel>();

            this.viewModels.Add(new PointViewModel(new Point(3, 3)));

            this.viewModels.Add(new PointViewModel(new Point(currentSize_X - 8, currentSize_Y - 8)));

            this.viewModels.Add(new PointViewModel(new Point(1, currentSize_Y - 8)));

            this.viewModels.Add(new PointViewModel(new Point(currentSize_X - 8, 1)));

        }



        public List<PointViewModel> Models

        {

            get { return this.viewModels; }

        }

    }

    public class PointViewModel

    {

        private Point point;

        public PointViewModel(Point point)

        {

            this.point = point;

        }



        public Double X { get { return point.X; } }

        public Double Y { get { return point.Y; } }
    }

}
