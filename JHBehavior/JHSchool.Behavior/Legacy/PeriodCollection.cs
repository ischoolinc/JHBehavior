using System;
using System.Collections.Generic;
using System.Text;

namespace JHSchool.Behavior.Legacy
{
    public class PeriodCollection
    {
        public PeriodCollection()
        {
            _periodList = new List<PeriodInfo>();
        }

        private List<PeriodInfo> _periodList;

        public List<PeriodInfo> Items
        {
            get { return _periodList; }
        }

        public List<PeriodInfo> GetSortedList()
        {
            _periodList.Sort(new PeriodComparer());
            return _periodList;
        }
    }
}
