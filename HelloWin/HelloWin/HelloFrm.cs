using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HelloWin
{
    public partial class HelloFrm : Form
    {
        public HelloFrm()
        {
            InitializeComponent();
        }

        private void HelloFrm_Load(object sender, EventArgs e)
        {
            this.Text = "我的第一个C#程序";
            Label lblShow = new Label();
            lblShow.Location = new Point(50, 60);
            lblShow.AutoSize = true;
            lblShow.Text = "欢迎你的使用";
            this.Controls.Add(lblShow);

        }
    }
}
