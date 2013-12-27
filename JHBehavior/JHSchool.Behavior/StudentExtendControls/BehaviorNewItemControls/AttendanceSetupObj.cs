using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace JHSchool.Behavior.StudentExtendControls
{
    class AttendanceSetupObj
    {
        public AttendanceSetupObj()
        {

        }

        public AttendanceSetupObj(XmlElement xml)
        {
            PeriodType = xml.GetAttribute("PeriodType");
            Name = xml.GetAttribute("Name");

            int CountInt;
            if (int.TryParse(xml.GetAttribute("Count"), out CountInt))
            {
                Count = CountInt;
            }
            else
            {
                Count = 0;
            }

            PeritodTypeName = xml.GetAttribute("PeriodType") + xml.GetAttribute("Name");
        }
        /// <summary>
        /// 類型
        /// </summary>
        public string PeriodType { get; set; }
        /// <summary>
        /// 缺曠名稱
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 統計數字
        /// </summary>
        public int Count { get; set; }
        /// <summary>
        /// 識別名稱
        /// </summary>
        public string PeritodTypeName { get; set; }

    }
}
