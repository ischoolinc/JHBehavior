using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JHSchool.Behavior.Report
{
    public static class tool
    {
        public static int SortPeriod(K12.Data.PeriodMappingInfo info1,K12.Data.PeriodMappingInfo info2)
        {
            return info1.Sort.CompareTo(info2.Sort);
        }
    }
}
