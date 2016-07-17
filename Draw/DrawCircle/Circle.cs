using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Shapes;

namespace DrawCircle
{
    class Circle : DrawingSharp
    {
        public Circle(int sidelength)
            : base(sidelength)
        { 
        }
        public override void Draw(System.Windows.Controls.Canvas canvas)
        {
            if (this.shape != null)
            {
                canvas.Children.Remove(this.shape);
            }
            else
            {
                this.shape = new Ellipse();

            }
            base.Draw(canvas);
        }
    }
}
