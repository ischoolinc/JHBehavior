using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JHSchool.Behavior.StuAdminExtendControls.BehaviorStatistics
{
    public class PeriodMap
    {
        #region new一個自定義的對照表

        private Dictionary<string, string> _data;

        public PeriodMap(Dictionary<string, string> data) //建構子
        {
            _data = data;
        }

        public string GetPeriodType(string period) //取得對照表
        {
            if (_data.ContainsKey(period))
                return _data[period];
            else
                return "{未定義}";
        }

        #endregion
    }
}
