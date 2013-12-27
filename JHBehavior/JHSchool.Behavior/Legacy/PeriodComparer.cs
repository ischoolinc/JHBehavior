using System;
using System.Collections.Generic;
using System.Text;

namespace JHSchool.Behavior.Legacy
{
    public class PeriodComparer : IComparer<PeriodInfo>
    {
        #region IComparer<PeriodInfo> 成員

        public int Compare(PeriodInfo x, PeriodInfo y)
        {
            if (x.Sort == y.Sort) return 0;
            else if (x.Sort > y.Sort) return 1;
            else return -1;
        }

        #endregion
    }
}
