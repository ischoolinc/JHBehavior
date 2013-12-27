using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace JHSchool.Behavior.StuAdminExtendControls
{
    public class AbsenceInfo
    {
        public AbsenceInfo()
        { }

        public AbsenceInfo(XmlElement element)
        {
            _name = element.GetAttribute("Name");
            _hotkey = element.GetAttribute("HotKey");
            _abbreviation = element.GetAttribute("Abbreviation");
            _subtract = element.GetAttribute("Subtract");
            _aggregated = element.GetAttribute("Aggregated");
            bool b;
            bool.TryParse(element.GetAttribute("Noabsence"), out b);
            _noabsence = b;
        }

        #region 假別名稱
        private string _name;

        /// <summary>
        /// 假別名稱
        /// </summary>
        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }
        #endregion
        #region 假別熱鍵
        private string _hotkey;

        /// <summary>
        ///  假別熱鍵
        /// </summary> 
        public string Hotkey
        {
            get { return _hotkey; }
            set { _hotkey = value; }
        }
        #endregion
        #region 假別縮寫
        private string _abbreviation;

        /// <summary>
        /// 假別縮寫
        /// </summary>
        public string Abbreviation
        {
            get { return _abbreviation; }
            set { _abbreviation = value; }
        }
        #endregion
        #region 扣分
        private string _subtract;
        /// <summary>
        /// 扣分
        /// </summary>
        public string Subtract
        {
            get { return _subtract; }
        }
        #endregion
        #region 累計缺曠節次
        private string _aggregated;
        /// <summary>
        /// 累計缺曠節次
        /// </summary>
        public string Aggregated
        {
            get { return _aggregated; }
        }
        #endregion
        #region 不影響全勤
        private bool _noabsence;
        /// <summary>
        /// 不影響全勤
        /// </summary>
        public bool Noabsence
        {
            get { return _noabsence; }
        }
        #endregion

        public static AbsenceInfo Empty
        {
            get
            {
                AbsenceInfo info = new AbsenceInfo();
                info.Name = string.Empty;
                info.Abbreviation = string.Empty;
                info.Hotkey = string.Empty;
                return info;
            }
        }

        public AbsenceInfo Clone()
        {
            AbsenceInfo info = new AbsenceInfo();
            info.Abbreviation = _abbreviation;
            info.Hotkey = _hotkey;
            info.Name = _name;
            return info;
        }
    }
}
