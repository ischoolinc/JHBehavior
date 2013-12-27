using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JHSchool.Behavior
{
    /// <summary>
    /// 由 Attendance 類別所提供的 Extend Method
    /// </summary>
    public static class Attendance_ExtendMethod
    {
        /// <summary>
        /// 取得學生缺曠資料。
        /// </summary>
        public static List<AttendanceRecord> GetAttendances(this StudentRecord student)
        {
            return Attendance.Instance[student.ID];
        }

        /// <summary>
        /// 批次同步學生缺曠資料。
        /// </summary>
        /// <param name="students"></param>
        public static void SyncAttendanceCache(this IEnumerable<StudentRecord> students)
        {
            Attendance.Instance.SyncDataBackground(students.AsKeyList());
        }
    }
}
