using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileAutomation_1._0.Model
{
    public class CursorPoint
    {
        public double X { get; set; }
        public double Y { get; set; }

        public CursorPoint(double x, double y)
        {
            X = x;
            Y = y;
        }
    }
}
