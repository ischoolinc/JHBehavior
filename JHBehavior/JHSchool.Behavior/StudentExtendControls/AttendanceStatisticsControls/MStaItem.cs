using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JHSchool.Behavior.StudentExtendControls.AttendanceStatisticsControls
{
    public class MStaItem
    {
        public string SchoolYear { get; set; }
        public string Semester { get; set; }
        public string MeritType { get; set; }
        public int Count { get; set; }
        //public Dictionary<string,int> MeritMapping{ get; set; }           
    }
}
