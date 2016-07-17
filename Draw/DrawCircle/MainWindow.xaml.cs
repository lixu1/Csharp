using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace DrawCircle
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void drawingCanvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            label1.Content = "画正方形";
            Point mouseLocation = e.GetPosition(this.drawingCanvas);
            Square mySquare = new Square(100);
            mySquare.SetLocation((int)mouseLocation.X, (int)mouseLocation.Y);
            mySquare.Draw(drawingCanvas);
            mySquare.SetColor(Colors.BlueViolet);
        }

        private void drawingCanvas_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            label1.Content = "画圆形";
            Point mouseLocation = e.GetPosition(this.drawingCanvas);
            Circle myCircle = new Circle(100);
            myCircle.SetLocation((int)mouseLocation.X, (int)mouseLocation.Y);
            myCircle.Draw(drawingCanvas);
            myCircle.SetColor(Colors.HotPink);

        }
    }
}
