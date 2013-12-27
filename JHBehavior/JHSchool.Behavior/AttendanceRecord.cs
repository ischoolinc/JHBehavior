using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using Framework;

namespace JHSchool.Behavior
{
    //<Attendance ID="1718753">
    //    <Detail>
    //        <Attendance>
    //            <Period AbsenceType="病假1" AttendanceType="一般">二</Period>
    //            <Period AbsenceType="病假1" AttendanceType="一般">三</Period>
    //            <Period AbsenceType="病假1" AttendanceType="一般">四</Period>
    //            <Period AbsenceType="病假1" AttendanceType="一般">五</Period>
    //            <Period AbsenceType="病假1" AttendanceType="一般">六</Period>
    //            <Period AbsenceType="病假1" AttendanceType="一般">七</Period>
    //            <Period AbsenceType="病假1" AttendanceType="一般">一</Period>
    //        </Attendance>
    //    </Detail>
    //    <Semester>1</Semester>
    //    <OccurDate>2005-09-02 00:00:00</OccurDate>
    //    <RefStudentID>168920</RefStudentID>
    //    <SchoolYear>94</SchoolYear>
    //</Attendance>

    /// <summary>
    /// 學生缺曠紀錄
    /// </summary>
    public class AttendanceRecord
    {
        public AttendanceRecord(XmlElement element)
        {
            XmlHelper helper = new XmlHelper(element);

            RefStudentID = helper.GetString("RefStudentID");
            ID = helper.GetString("@ID");
            SchoolYear = helper.GetString("SchoolYear");
            Semester = helper.GetString("Semester");
            OccurDate = helper.GetDateString("OccurDate");

            PeriodDetail = new List<AttendancePeriod>();
            foreach(XmlElement each in helper.GetElements("Detail/Attendance/Period"))
            {
                PeriodDetail.Add(new AttendancePeriod(each));
            }
        }

        public string RefStudentID { get; private set; }    //學生編號
        public string ID { get; private set; }                          //記錄編號
        public string SchoolYear { get; private set; }          //學年度
        public string Semester { get; private set; }             //學期
        public string OccurDate { get; private set; }           //缺曠日期
        public string DayOfWeek
        {
            get 
            {
                DateTime date;
                if (DateTime.TryParse(OccurDate, out date))
                    return DayOfWeekInChinese(date.DayOfWeek);

                return "";
            }
        }
        public List<AttendancePeriod> PeriodDetail { get; private set; }        //記錄內容

        private string DayOfWeekInChinese(DayOfWeek day)
        {
            switch (day)
            {
                case System.DayOfWeek.Monday:
                    return "一";
                case System.DayOfWeek.Tuesday:
                    return "二";
                case System.DayOfWeek.Wednesday:
                    return "三";
                case System.DayOfWeek.Thursday:
                    return "四";
                case System.DayOfWeek.Friday:
                    return "五";
                case System.DayOfWeek.Saturday:
                    return "六";
                default:
                    return "日";
            }
        }
    }

    //<Period AbsenceType="病假1" AttendanceType="一般">二</Period>
    //<Period AbsenceType="病假1" AttendanceType="一般">三</Period>

    /// <summary>
    /// 缺曠紀錄內容
    /// </summary>
    public class AttendancePeriod
    {
        public AttendancePeriod() { }

        internal AttendancePeriod(XmlElement element)
        {
            Period = element.InnerText;
            AbsenceType = element.Attributes["AbsenceType"].InnerText;
            //AttendanceType = element.Attributes["AttendanceType"].InnerText;
        }

        public string Period { get; set; }                       //節次
        public string AbsenceType { get; set; }            //Absence Type
        //public string AttendanceType { get; set; }       //Attendance Type
    }
}
