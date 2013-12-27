using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JHSchool.Behavior.StudentExtendControls.AttendanceStatisticsControls
{
    public class AbsItem
    {
        public string SchoolYear { get; set; }
        public string Semester { get; set; }
        public string Name { get; set; }
        public string PeriodType { get; set; }
        public int Count { get; set; }        
    }
}
