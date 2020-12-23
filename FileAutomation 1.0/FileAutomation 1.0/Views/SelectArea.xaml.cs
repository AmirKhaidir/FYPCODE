using FileAutomation_1._0.Model;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace FileAutomation_1._0.Views
{
    /// <summary>
    /// Interaction logic for SelectArea.xaml
    /// </summary>
    public partial class SelectArea : Window
    {
        //Moving window by click-drag on a control https://stackoverflow.com/a/13477624/5260872
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HTCAPTION = 0x2;

        [DllImport("User32.dll")]
        public static extern bool ReleaseCapture();

        [DllImport("User32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        public ScreenShot shotData { get; set; }
        public SelectArea()
        {
            InitializeComponent();

            System.Windows.Application curApp = System.Windows.Application.Current;
            Window mainWindow = curApp.MainWindow;
            mainWindow.WindowState = WindowState.Minimized;
            this.Activate();
            this.Topmost = true;
            this.Opacity = .5D; //Make trasparent

        }


        private void canvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (Mouse.LeftButton == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            int w = (int)canvas.ActualWidth;
            int h = (int)canvas.ActualHeight;

            System.Drawing.Size sz = new System.Drawing.Size((int)ActualWidth, (int)ActualHeight);
            System.Drawing.Point loc = new System.Drawing.Point((int)Left, (int)Top);

            shotData = new ScreenShot
            {
                width = w,
                height = h,
                x = loc.X,
                y = loc.Y,
                size = sz
            };

            this.Close();
            //System.Windows.Application curApp = System.Windows.Application.Current;
            //Window mainWindow = curApp.MainWindow;

            //mainWindow.Show();

            //MainWindow main = new MainWindow(loc.X, loc.Y, w, h, sz);

            //mainWindow.Show();
        }
    }
}
