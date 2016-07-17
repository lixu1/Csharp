using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Shapes;
namespace DrawCircle
{
    class Square:DrawingSharp
    {
        public Square(int sidelength)
            : base(sidelength)
        { 
        }
        public override void Draw(System.Windows.Controls.Canvas canvas)
        {
            if (this.shape != null)
            {
                canvas.Children.Remove(this.shape);
            }
            this.shape = new Rectangle();
            base.Draw(canvas);
        }
    }
}
