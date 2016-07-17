using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Drawing;
using System.Collections;
namespace GDI_example
{
    /// <summary>
    /// 生成条形码图形、条形码代码字符，条形码打印等处理的类；
    /// 开发人员：###；
    /// 开发时间：####；
    /// </summary>
    public  class BarCode
    {
        /*
         * 编码规则：
           39码总共可表示的字符范围包含有：0~9的数字，A~Z的英文字母，以及“＋”、“－”、“＊”、“／”、“％”、“＄”、“．”
           等特殊符号，再加上空格符“ ”，共计44组编码，每一个字符编码都是按照5条四空，其中有3个宽单元，即藉由九条不同排列的线条编码而得。
         * 可分为四类：
　            粗黑线：11
　            细黑线：1
　            粗白线：00
　            细白线：0
　            而且规范中的规定了每个字符所对应的编码对应表，所以生成条码实质上只是根据每个字符的编码画出宽度不同的线而已。
         */       
        public static int WidthL3 = 3;    //粗线和宽间隙宽度
        public static int WidthL1 = 1;    //细线和窄间隙宽度
        public static int LineX = 10;   //条码起始坐标
        public static int LineHeight = 75;
        //画布大小定义
        private static int Height = 100;
        private static int Width = 20;
        #region code39
        private static String alphabet39 = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%*";
        private static String[] coded39Char = 
		{
			/* 0 */ "000110100", 
			/* 1 */ "100100001", 
			/* 2 */ "001100001", 
			/* 3 */ "101100000",
			/* 4 */ "000110001", 
			/* 5 */ "100110000", 
			/* 6 */ "001110000", 
			/* 7 */ "000100101",
			/* 8 */ "100100100", 
			/* 9 */ "001100100", 
			/* A */ "100001001", 
			/* B */ "001001001",
			/* C */ "101001000", 
			/* D */ "000011001", 
			/* E */ "100011000", 
			/* F */ "001011000",
			/* G */ "000001101", 
			/* H */ "100001100", 
			/* I */ "001001100", 
			/* J */ "000011100",
			/* K */ "100000011", 
			/* L */ "001000011", 
			/* M */ "101000010", 
			/* N */ "000010011",
			/* O */ "100010010", 
			/* P */ "001010010", 
			/* Q */ "000000111", 
			/* R */ "100000110",
			/* S */ "001000110", 
			/* T */ "000010110", 
			/* U */ "110000001", 
			/* V */ "011000001",
			/* W */ "111000000", 
			/* X */ "010010001", 
			/* Y */ "110010000", 
			/* Z */ "011010000",
			/* - */ "010000101", 
			/* . */ "110000100", 
			/*' '*/ "011000100",
			/* $ */ "010101000",
			/* / */ "010100010", 
			/* + */ "010001010", 
			/* % */ "000101010", 
			/* * */ "010010100" 
		};
        #endregion
   
        /// <summary>
        /// 对输入字符进行Code39编码
        /// </summary>
        /// <param name="Code">要编码的字符串</param>
        /// <returns>编码后的字符串</returns>
        private static string Encode39(string Code)
        {
            Code = Code.ToUpper();
             //标准化输入字符，增加起始标记
              if (Code.Substring(0,1) != "*") 
              {
               Code = "*" + Code;
              }
              if (Code.Substring(Code.Length-1,1)!="*")
              {
               Code = Code + "*";
              }
             //转换成39编码
             string Code39="";
             for (int i = 0; i < Code.Length; i++) 
             {
                 Code39 = Code39 + coded39Char[alphabet39.IndexOf(Code[i])].ToString();
             }
             return Code39;
        }
        /// <summary>
        /// 动态设置图片的宽度，根据数输入字符的长度
        /// </summary>
        /// <param name="Code39">编码后的字符串</param>
        /// <returns>图片宽度</returns>
        private static int GetBarCodeWidth(string Code39)
        {
            int w=LineX*2;
            for (int i = 0; i < Code39.Length; i++)
            {
                int sw = 0;
                //0 
                if (Code39[i].ToString() == "0")
                {
                    sw = WidthL1;
                }
                //1
                else
                {
                    sw = WidthL3;
                }
                w += sw;
            }
            return w;
        }
        /// <summary>
        /// 获取条形码图片，字符串要求："0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%*"；
        /// 不能有其他的非法字符，暂时没有做相关的错误处理；
        /// 如果有非法字符，程序会报错
        /// </summary>
        /// <param name="Code">要生产条形码图标的字符串</param>
        /// <returns>条形码图片</returns>
        public static  Bitmap GetDrawBarImage(string Code)
        {
            string code39 = Encode39(Code);
            //设置宽度
            Width = GetBarCodeWidth(code39);
            Code = Code.ToUpper();
            //generate image graphics
            Bitmap image = new Bitmap(Width, Height);
            Graphics g = Graphics.FromImage(image);
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            try
            {
                g.Clear(Color.White);
                //style 样式
                Pen penB = new Pen(Color.Black);
                Pen penW = new Pen(Color.White);
                Font f = new Font("ARial", 9);
                //画线
                int CL = LineX;
                for (int i = 0; i < code39.Length; i++)
                {
                    int sw = 0;
                    //0 
                    if (code39[i].ToString() == "0")
                    {
                       
                        sw = WidthL1;
                    }
                    //1
                    else
                    {
                        sw = WidthL3;
                    }
                    Rectangle rect = new Rectangle(CL, 5, sw, LineHeight);
                    g.FillRectangle(i % 2 == 0 ? new SolidBrush(Color.Black) : new SolidBrush(Color.White), rect);
                    CL += sw;
                }
                //绘制标题
                Point p = new Point(Width/2 - Convert.ToInt32( Code.Length * 4.5), Height - 18);
                g.DrawString(Code, f, new SolidBrush(Color.Black), p);
                //return
                return image;
            }
            finally
            {
                g.Dispose();
            }
           
    }
    }
}
