using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace simpleGrammar
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        enum MyEnum { a = 101, b, c, d = 201, e, f };

        struct Student
        {
            public int no;
            public string name;
            public char sex;
            public int score;
            public string Answer()
            {
                string result = "该学生的信息如下：";
                result += "\n学号：" + no;
                result += "\n姓名：" + name;
                result += "\n性别：" + sex;
                result += "\n成绩:" + score;
                return result;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            int a = 12, b = 15, c, d;
            c = a + b;
            d = c * b;
            lblShow.Text = "变量c的值为：" + c;
            lblShow.Text += "\n变量d的值为：" + d;
            MyEnum x = MyEnum.f;
            MyEnum y = (MyEnum)202;
            string result = "\n枚举型x的值为" + (int)x;
            result += "\n枚举型y代表枚举元素" + y;
            lblShow.Text += result;

            result="\n现在时间是:"+DateTime.Now+"\n";
            lblShow.Text += result;

            Student s;
            s.no = 1001;
            s.name = "李明";
            s.sex = '男';
            s.score = 100;
            lblShow.Text += s.Answer();


            int i = 1, j = 1, p, q;
            p = (i++) + (i++) + (i++);
            q = (++j) + (++j) + (++j);
            lblShow.Text += "\n变量i的值为" + i;
            lblShow.Text += "\n变量j的值为" + j;
            lblShow.Text += "\n变量p的值为" + p;
            lblShow.Text += "\n变量q的值为" + q;


            int[] a1 = { 5, 1, 2, 4, 3 };
            int[] b1 = new int[5];
            Array.Copy(a1, b1, 5);
            Array.Clear(a1, 0, 5);
            lblShow.Text += "\n" + b1[0] + " " + b1[1] + " " + b1[2] + " " + b1[3] + " " + b1[4] + "\n";
            Array.Sort(b1);
            lblShow.Text += b1[0] + " " + b1[1] + " " + b1[2] + " " + b1[3] + " " + b1[4] + "\n";
            Array.Reverse(b1);
            lblShow.Text += b1[0] + " " + b1[1] + " " + b1[2] + " " + b1[3] + " " + b1[4] + "\n";

            int[,] a3 = new int[2, 3] { { 1, 2, 3 }, { 4, 5, 6 } };
            int[][] b3 = new int[2][];
            b3[0] = new int[3] { 1, 2, 3 };
            b3[1] = new int[4] { 4, 5, 6, 7 };
            lblShow.Text += "\na3是一个二维数组，均为整数值";
            lblShow.Text += "\nb3是一个一位数组，均为子数组";
            lblShow.Text += "\na3[0,0]=" + a3[0, 0];
            lblShow.Text += "\nb3[0][0]=" + b3[0][0];
        }
    }
}
