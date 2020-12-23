using FileAutomation_1._0.Library;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileAutomation_1._0.Model
{
    public class MouseEvent
    {
        public CursorPoint Location { get; set; }
        public MouseHook.MouseEvents Action { get; set; }
        public int Value { get; set; }
    }
}
