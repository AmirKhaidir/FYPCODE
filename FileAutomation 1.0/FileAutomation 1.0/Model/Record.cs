using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileAutomation_1._0.Model
{
    public class Record
    {
        public int Id { get; set; }
        public string Content { get; set; }
        public MouseEvent EventMouse { get; set; }
        public KeyboardEvent EventKey { get; set; }
        public string ElementName { get; set; }
        public string Type { get; set; }
        public int WaitMs { get; set; }
    }
}
