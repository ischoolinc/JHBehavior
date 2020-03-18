using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Transfer_Excel
{
    /// <summary>
    /// 代表一筆要列印的資料
    /// </summary>
    class DisciplineEntity
    {
        public String StudentID { get; set; }
        public String DisciplineID { get; set; }
        public DateTime OccurDate { get; set; }
        public String Reason { get; set; }

        public String Remark { get; set; }

        public String RegisterDate { get; set; }
        public String MeritFlag { get; set; }
        public String ACount { get; set; }
        public String BCount { get; set; }
        public String CCount { get; set; }

        public String Clear { get; set; }
        public String ClearDate { get; set; }
        public String ClearReason { get; set; }



        public DisciplineEntity(XmlElement elm)
        {
            StudentID = elm.SelectSingleNode("RefStudentID").InnerText;
            OccurDate = DateTime.Parse(elm.SelectSingleNode("OccurDate").InnerText);
            DisciplineID = elm.GetAttribute("ID");
            
            Reason = elm.SelectSingleNode("Reason").InnerText;
            Remark = elm.SelectSingleNode("Remark").InnerText;
            RegisterDate = elm.SelectSingleNode("RegisterDate").InnerText;
            MeritFlag = elm.SelectSingleNode("MeritFlag").InnerText;

            parseABCValues(elm, MeritFlag);


        }

        private void parseABCValues(XmlElement elm, String meritFlag )
        {
            if (meritFlag == "1")
            {
                XmlElement meritElement = (XmlElement)elm.SelectSingleNode("Detail/Discipline/Merit");
                if (meritElement == null) 
                    return;
                ACount = meritElement.GetAttribute("A");
                BCount = meritElement.GetAttribute("B");
                CCount = meritElement.GetAttribute("C");

            }
            else if (meritFlag == "0")
            {
                XmlElement demeritElement = (XmlElement)elm.SelectSingleNode("Detail/Discipline/Demerit");
                if (demeritElement == null) 
                    return;
                ACount = demeritElement.GetAttribute("A");
                BCount = demeritElement.GetAttribute("B");
                CCount = demeritElement.GetAttribute("C");

                Clear = demeritElement.GetAttribute("Cleared");
                ClearDate = demeritElement.GetAttribute("ClearDate");
                ClearReason = demeritElement.GetAttribute("ClearReason");

            }
        }
    }
}
