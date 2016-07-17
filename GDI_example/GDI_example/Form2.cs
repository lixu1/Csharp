using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Drawing.Imaging;

namespace GDI_example
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
           
        }
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofdlg = new OpenFileDialog();
            ofdlg.Filter = "BMP File(*.bmp)|*.bmp|*.jpg|*.gif";
            if (ofdlg.ShowDialog() == DialogResult.OK)
            {
                Bitmap image = new Bitmap(ofdlg.FileName);
                pictureBox1.Image = image;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string str;
            Bitmap box1 = new Bitmap(pictureBox1.Image);
            SaveFileDialog sfdlg = new SaveFileDialog();
            sfdlg.Filter = "bmp文件(*.BMP)|*.BMP|ALL File(*.*)|*.*";
            sfdlg.ShowDialog();
            str = sfdlg.FileName;
            box1.Save(str);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string str;
            Bitmap box1 = new Bitmap(pictureBox1.Image);
            SaveFileDialog sfdlg = new SaveFileDialog();
            sfdlg.Filter = "bmp文件(*.jpeg)|*.jpeg|ALL File(*.*)|*.*";
            sfdlg.ShowDialog();
            str = sfdlg.FileName;
            box1.Save(str, System.Drawing.Imaging.ImageFormat.Jpeg);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(pictureBox1.Image);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            IDataObject idata = Clipboard.GetDataObject();
            if (idata.GetDataPresent(DataFormats.Bitmap))
            {
                pictureBox2.Image = (Bitmap)idata.GetData(DataFormats.Bitmap);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Color c = new Color();
            Bitmap box1 = new Bitmap(pictureBox1.Image);
            Bitmap box2 = new Bitmap(pictureBox1.Image);
            int r, g, b, size, k1, k2, xres, yres, i, j;
            xres = pictureBox1.Image.Width;
            yres = pictureBox1.Image.Height;
            size = 4;
            for (i = 0; i <= xres - 1; i += size)
            {
                for (j = 0; j <= yres - 1; j += size)
                {
                    c = box1.GetPixel(i, j);
                    r = c.R;
                    g = c.G;
                    b = c.B;
                    Color cc = Color.FromArgb(r, g, b);
                    for (k1 = 0; k1 <= size - 1; k1++)
                    {
                        for (k2 = 0; k2 <= size - 1; k2++)
                        {
                            if (i + k1 < pictureBox1.Image.Width)
                                box2.SetPixel(i + k1, j + k2, cc);
                        }
                    }
                }
            }
            pictureBox2.Refresh();
            pictureBox2.Image = box2;

        }

        private void button7_Click(object sender, EventArgs e)
        {
            Color c=new Color();
            Bitmap b = new Bitmap(pictureBox1.Image);
            Bitmap b1 = new Bitmap(pictureBox1.Image);
            int rr, gg, bb, cc,i,j;
            for (i = 0; i < pictureBox1.Image.Width; i++)
            {
                for (j = 0; j < pictureBox1.Image.Height; j++)
                {
                    c = b.GetPixel(i, j);
                    rr = c.R;
                    gg = c.G;
                    bb = c.B;
                    cc = (int)((rr + gg + bb) / 3);
                    if (cc < 0) cc = 0;
                    if (cc > 255) cc = 255;
                    Color c1 = Color.FromArgb(cc, cc, cc);
                    b1.SetPixel(i, j, c1);
                }
                
            }

            pictureBox2.Refresh();
            pictureBox2.Image = b1;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            //Color c = new Color();
            Bitmap box1 = new Bitmap(pictureBox1.Image);
            Bitmap box2 = new Bitmap(pictureBox1.Image);
            int rr, x, m, lev, wid;
            int[] lut = new int[256];
            int[, ,] pic = new int[600, 600, 3];
            double dm;
            lev = 80;
            wid = 100;
            for (x = 0; x < 256; x++)
            {
                lut[x] = 255;
            }
            for (x = lev; x < lev + wid; x++)
            {
                dm = ((double)(x - lev) / (double)wid) * 255f;
                lut[x] = (int)dm;
            }
            for (int i = 0; i < pictureBox1.Image.Width - 1; i++)
            {
                for (int j = 0; j < pictureBox1.Image.Height; j++)
                {
                    m = pic[i, j, 0];
                    rr = lut[m];
                    Color c1 = Color.FromArgb(rr, rr, rr);
                    box2.SetPixel(i, j, c1);
                }
                pictureBox2.Refresh();
                pictureBox2.Image = box2;
            }

        }


    }
}
