using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Behavior.MeritDemeritStatistics
{
    public class MeritDemeritObj
    {
        public Dictionary<string, int> 一年級男生獎懲 { get; set; }
        public Dictionary<string, int> 二年級男生獎懲 { get; set; }
        public Dictionary<string, int> 三年級男生獎懲 { get; set; }

        public Dictionary<string, int> 一年級女生獎懲 { get; set; }
        public Dictionary<string, int> 二年級女生獎懲 { get; set; }
        public Dictionary<string, int> 三年級女生獎懲 { get; set; }

        public Dictionary<string, int> 男生獎懲 { get; set; }
        public Dictionary<string, int> 女生獎懲 { get; set; }

        public Dictionary<string, int> 總獎懲 { get; set; }

        public MeritDemeritObj()
        {
            一年級男生獎懲 = new Dictionary<string, int>();
            二年級男生獎懲 = new Dictionary<string, int>();
            三年級男生獎懲 = new Dictionary<string, int>();

            一年級男生獎懲 = new Dictionary<string, int>();
            二年級男生獎懲 = new Dictionary<string, int>();
            三年級男生獎懲 = new Dictionary<string, int>();

            男生獎懲 = new Dictionary<string, int>();
            女生獎懲 = new Dictionary<string, int>();
            總獎懲 = new Dictionary<string, int>();
        }
    }
}
