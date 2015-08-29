using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SearchStringBuilder
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        StringBuilder sb = new StringBuilder();

        private void btnAdd_Click(object sender, EventArgs e)
        {
            sb.Append(txtSource.Text);
            lblShow.Text = "字符串\"" + sb.ToString() + "\"的长度为" + sb.Length;

        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string s = sb.ToString();
            int pos = s.IndexOf(txtSearch.Text);
            lblShow.Text += "\n目标起始索引值为" + pos;

        }
    }
}
