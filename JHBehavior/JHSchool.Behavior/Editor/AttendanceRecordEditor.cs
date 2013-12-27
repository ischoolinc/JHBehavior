using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Framework;
using System.Xml;
using JHSchool.Editor;

namespace JHSchool.Behavior.Editor
{
    public class AttendanceRecordEditor
    {
        /// <summary>
        /// Constructor，為修改模式使用
        /// </summary>
        /// <param name="demeritRecord"></param>
        internal AttendanceRecordEditor(AttendanceRecord updateRecord)
        {
            UpdateRecord = updateRecord;
            
            Remove = false;
            RefStudentID = UpdateRecord.RefStudentID;
            ID = UpdateRecord.ID;

            SchoolYear = updateRecord.SchoolYear;
            Semester = updateRecord.Semester;
            OccurDate = updateRecord.OccurDate;
            PeriodDetail = updateRecord.PeriodDetail;
        }

        /// <summary>
        /// Constructor ，為新增模式使用。
        /// </summary>
        /// <param name="studentRecord"></param>
        internal AttendanceRecordEditor(string PrimaryKey)
        {
            Remove = false;
            RefStudentID = PrimaryKey;
        }

        /// <summary>
        /// 取得編輯狀態
        /// </summary>
        public EditorStatus EditorStatus
        {
            get
            {
                if (UpdateRecord == null)
                {
                    if (!Remove)
                        return EditorStatus.Insert;
                    else
                        return EditorStatus.NoChanged;
                }
                else
                {
                    if (Remove)
                        return  EditorStatus.Delete;

                    else if (UpdateRecord.SchoolYear != SchoolYear ||
                        UpdateRecord.Semester != Semester ||
                        UpdateRecord.OccurDate != OccurDate ||
                        UpdateRecord.PeriodDetail != PeriodDetail)
                    {
                        return EditorStatus.Update;
                    }
                }

                return EditorStatus.NoChanged;
            }
        }

        public void Save()
        {
            if (EditorStatus != EditorStatus.NoChanged)
                Feature.EditAttendance.SaveAttendanceRecordEditor(new AttendanceRecordEditor[] { this });
        }

        #region Fields

        public bool Remove { get; set; }

        internal string RefStudentID { get; private set; }
        public string ID { get; private set; }

        public string SchoolYear { get; set; }
        public string Semester { get; set; }
        public string OccurDate { get; set; }
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
        public List<AttendancePeriod> PeriodDetail { get; set; }

        #endregion

        internal AttendanceRecord UpdateRecord { get; private set; }

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

    public static class AttendanceRecordEditor_ExtendMethods
    {
        public static AttendanceRecordEditor GetEditor(this AttendanceRecord record)
        {
            return new AttendanceRecordEditor(record);
        }
    }
}
