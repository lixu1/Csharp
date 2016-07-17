using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GDI_example
{
    //根据统计数据，输出各种统计图形，包括饼状图，曲线分析图，柱状图，
    //多组数据曲线分析图。
    //图形统一大小600*420
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
            pictureBox1.Image = CountImage.GetPieImage("条形图", new String[] { "一月", "二月", "三月" }, new decimal[] { 5, 10,11});
        }
    }
}