using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JHSchool.Behavior.StuAdminExtendControls.BehaviorStatistics
{
    internal class DisciplineItem
    {
        #region 獎懲物件
        public DisciplineItem()
        {
            A = 0;
            B = 0;
            C = 0;
        }

        public int A { get; set; }

        public int B { get; set; }

        public int C { get; set; }

        #endregion
    }
}
