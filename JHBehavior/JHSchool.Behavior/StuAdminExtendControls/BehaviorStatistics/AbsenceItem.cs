using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JHSchool.Behavior.StuAdminExtendControls.BehaviorStatistics
{
    internal class AbsenceItem
    {
        #region 自定義之缺曠 單一節次 物件
        public AbsenceItem()
        {
            Count = 0;
        }

        public void Add()
        {
            Count++;
        }

        public int Count { get; set; }

        public string Name { get; set; }

        public string Type { get; set; }
        #endregion
    }
}
