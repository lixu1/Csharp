using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WeekDay
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        enum WeekDay { 星期天, 星期一, 星期二, 星期三, 星期四, 星期五, 星期六 };


        private void button1_Click(object sender, EventArgs e)
        {
            DateTime dt = Convert.ToDateTime(textBox1.Text);
            WeekDay wd = (WeekDay)dt.DayOfWeek;
            label2.Text = "这一天是" + wd + ".";
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
