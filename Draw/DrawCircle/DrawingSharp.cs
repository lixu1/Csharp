using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Controls;

namespace DrawCircle
{
    abstract class DrawingSharp
    {
        protected int size;
        protected int locx = 0, locy = 0;
        protected Shape shape = null;
        public DrawingSharp(int size)
        {
            this.size = size;
        }
        public void SetLocation(int xCoord, int yCoord)
        {
            this.locx = xCoord;
            this.locy = yCoord;
        }
        public void SetColor(Color color)
        {
            if(shape!=null)
            {
                SolidColorBrush brush = new SolidColorBrush(color);
                shape.Fill=brush;
            }
        }
        public virtual void Draw(Canvas canvas)
        {
            if (this.shape == null)
            {
                throw new ApplicationException("Shape is null");
            }
            this.shape.Height = this.size;
            this.shape.Width = this.size;
           Canvas.SetTop(this.shape, this.locy);
           Canvas.SetLeft(this.shape, this.locx);
            canvas.Children.Add(shape);
        }

    }
}
