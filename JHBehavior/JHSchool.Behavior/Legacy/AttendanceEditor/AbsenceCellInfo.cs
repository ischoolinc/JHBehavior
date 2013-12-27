using System;
using System.Collections.Generic;
using System.Text;

namespace JHSchool.Behavior.Legacy
{
    public class AbsenceCellInfo : ICellInfo<AbsenceInfo>
    {
        private AbsenceInfo _nowValue;
        private AbsenceInfo _absenceInfo;
        private bool _isdirty;
        private bool _valueChanged;

        public AbsenceCellInfo()
        {
            _nowValue = AbsenceInfo.Empty;
            _absenceInfo = _nowValue;
            _isdirty = false;
        }

        public AbsenceCellInfo(AbsenceInfo absenceInfo)
        {
            _nowValue = absenceInfo;
            _absenceInfo = absenceInfo;
            _isdirty = false;
        } 

        #region ICellInfo<AbsenceInfo> жин√

        public AbsenceInfo OriginValue
        {
            get { return _absenceInfo; }
        }

        public void SetValue(AbsenceInfo value)
        {
            if (value == null)
                value = AbsenceInfo.Empty;

            _valueChanged = (_nowValue.Name != value.Name);
                
            _nowValue = value;

            _isdirty = (_absenceInfo.Name != value.Name);

        }

        public bool IsValueChanged
        {
            get { return _valueChanged; }
        }

        public bool IsDirty
        {
            get { return _isdirty; }
        }

        public AbsenceInfo AbsenceInfo
        {
            get { return _nowValue; }
        }

        #endregion
    }
}
