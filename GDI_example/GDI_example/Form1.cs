using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Drawing.Design;
using System.Drawing.Text;

namespace GDI_example
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            //矩形
            Graphics g = e.Graphics;
            //Rectangle rect = new Rectangle(50, 30, 100, 100);
            //LinearGradientBrush lBrush = new LinearGradientBrush(rect,
            //    Color.Red, Color.Yellow, LinearGradientMode.BackwardDiagonal);
            //g.FillRectangle(lBrush, rect);

            //Pen pn = new Pen(Color.Blue);
            //Rectangle rect1 = new Rectangle(50, 30, 200, 100);
            ////弧线
            //g.DrawArc(pn, rect,0, 100);
            ////椭圆
            //g.DrawEllipse(pn, rect1);

            ////线段
            //Point pt1 = new Point(30, 30);
            //Point pt2 = new Point(110, 110);
            //g.DrawLine(pn, pt1, pt2);

            ////字体
            //Font fnt = new Font("Verdana", 16);
            //g.DrawString("GDI+ World",fnt,new SolidBrush(Color.Brown),54,50);

            ////填充路径
            ////g.FillRectangle(new SolidBrush(Color.White), ClientRectangle);
            //GraphicsPath path = new GraphicsPath(new Point[]
            //{
            //    new Point(40,140),new Point(275,200),new Point(105,225),new Point(190,300),
            //    new Point(50,350),new Point(20,180)
            //}, new byte[]
            //{
            //    (byte)PathPointType.Start,(byte)PathPointType.Bezier,(byte)PathPointType.Bezier,
            //    (byte)PathPointType.Bezier,(byte)PathPointType.Line,(byte)PathPointType.Line
            //});
            //PathGradientBrush pgb = new PathGradientBrush(path);
            //pgb.SurroundColors = new Color[]
            //{
            //    Color.Green,Color.Yellow,Color.Red,Color.Blue,Color.Orange,Color.White
            //};
            //g.FillPath(pgb, path);

            //单色画刷
            //SolidBrush sdBrush1 = new SolidBrush(Color.Red);
            //SolidBrush sdBrush2 = new SolidBrush(Color.Green);
            //SolidBrush sdBrush3 = new SolidBrush(Color.Blue);
            //g.FillEllipse(sdBrush2, 20, 40, 60, 70);
            //g.FillPie(sdBrush3, 0, 0, 200, 40, 0.0f, 300f);
            //PointF point1 = new PointF(50f, 250f);
            //PointF point2 = new PointF(100f, 25f);
            //PointF point3 = new PointF(150f, 40f);
            //PointF point4 = new PointF(250f, 50f);
            //PointF point5 = new PointF(300f, 100f);
            //PointF[] curvePoints = { point1, point2, point3, point4, point5 };
            //g.FillPolygon(sdBrush1, curvePoints);

            //HatchBrush画刷的使用
            //HatchBrush hBrush1 = new HatchBrush(HatchStyle.DiagonalCross,
            //    Color.Chocolate, Color.Red);
            //HatchBrush hBrush2 = new HatchBrush(HatchStyle.DashedHorizontal,
            //    Color.Green, Color.Black);
            //HatchBrush hBrush3 = new HatchBrush(HatchStyle.Weave,
            //    Color.BlueViolet, Color.Blue);
            //g.FillEllipse(hBrush1, 20, 80, 60, 20);
            //g.FillPie(hBrush3, 0, 0, 200, 40, 0f, 30f);
            //PointF point1 = new PointF(50f, 250f);
            //PointF point2 = new PointF(100f, 25f);
            //PointF point3 = new PointF(150f, 40f);
            //PointF point4 = new PointF(250f, 50f);
            //PointF point5 = new PointF(300f, 100f);
            //PointF[] curvePoints = { point1, point2, point3, point4, point5 };
            //g.FillPolygon(hBrush2, curvePoints);
           

            //纹理画刷
            Bitmap bitmap = new Bitmap("..\\..\\2.jpg");
            bitmap = new Bitmap(bitmap, this.ClientRectangle.Size);
            TextureBrush myBrush = new TextureBrush(bitmap);
            g.FillEllipse(myBrush, this.ClientRectangle);

            //渐变画刷
            //线性渐变
            //LinearGradientBrush myBrush = new LinearGradientBrush(
            //    this.ClientRectangle, Color.White, Color.Blue,
            //    LinearGradientMode.Vertical);
            //g.FillRectangle(myBrush, this.ClientRectangle);
            ////路径渐变
            //Point centerPoint = new Point(150, 100);
            //int R = 60;
            //GraphicsPath path = new GraphicsPath();
            //path.AddEllipse(centerPoint.X - R, centerPoint.Y - R, 2 * R, 2 * R);
            //PathGradientBrush brush = new PathGradientBrush(path);
            //brush.CenterPoint = centerPoint;
            //brush.CenterColor = Color.Red;
            //brush.SurroundColors = new Color[] { Color.Plum };
            //g.FillEllipse(brush, centerPoint.X - R, centerPoint.Y - R, 2 * R, 2 * R);
            //centerPoint = new Point(350, 100);
            //R = 20;
            //path = new GraphicsPath();
            //path.AddEllipse(centerPoint.X - R, centerPoint.Y - R, 2 * R, 2 * R);
            //path.AddEllipse(centerPoint.X - 2 * R, centerPoint.Y - 2 * R, 4 * R, 4 * R);
            //path.AddEllipse(centerPoint.X - 3 * R, centerPoint.Y - 3 * R, 6 * R, 6 * R);
            //brush = new PathGradientBrush(path);
            //brush.CenterPoint = centerPoint;
            //brush.CenterColor = Color.Red;
            //brush.SurroundColors = new Color[]{Color.Black,Color.Blue,
            //    Color.Green};
            //g.FillPath(brush, path);

        }

        

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
