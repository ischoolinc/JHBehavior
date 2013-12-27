using System;
using System.Collections.Generic;
using System.Text;

namespace JHSchool.Behavior.Legacy
{
    public class SemesterCellInfo:ICellInfo<string>
    {
        private string _semester;
        private bool _isdirty;

        public SemesterCellInfo(string semester)
        {
            _semester = semester;
            _isdirty = false;
        }
        #region ICellInfo<string> жин√

        public string OriginValue
        {
            get { return _semester; }
        }

        public void SetValue(string value)
        {
            if (value == null)
                value = string.Empty;

            _isdirty = (_semester != value);
        }

        public bool IsDirty
        {
            get { return _isdirty; }
        }

        #endregion 
    }
}
