using System;
using System.Collections.Generic;
using System.Text;
using JHSchool.Behavior.Feature;

namespace JHSchool.Behavior
{
    /// <summary>
    /// Cache manager of Attendance
    /// </summary>
    public class Attendance : Framework.CacheManager<List<AttendanceRecord>>
    {
        private static Attendance _Instance = null;

        private Attendance() { }

        public static Attendance Instance { get { if (_Instance == null)_Instance = new Attendance(); return _Instance; } }

        /// <summary>
        /// 取得所有的缺曠紀錄
        /// </summary>
        /// <returns></returns>
        protected override Dictionary<string, List<AttendanceRecord>> GetAllData()
        {
            Dictionary<string, List<AttendanceRecord>> oneToMany = new Dictionary<string, List<AttendanceRecord>>();

            foreach (AttendanceRecord each in QueryAttendance.GetAllAttendanceRecord())
            {
                if (!oneToMany.ContainsKey(each.RefStudentID))
                    oneToMany.Add(each.RefStudentID, new List<AttendanceRecord>());

                oneToMany[each.RefStudentID].Add(each);
            }

            return oneToMany;
        }

        /// <summary>
        /// 取得指定的學生的缺曠紀錄
        /// </summary>
        /// <param name="primaryKeys">學生ID的集合</param>
        /// <returns></returns>
        protected override Dictionary<string, List<AttendanceRecord>> GetData(IEnumerable<string> primaryKeys)
        {
            Dictionary<string, List<AttendanceRecord>> oneToMany = new Dictionary<string, List<AttendanceRecord>>();

            foreach (AttendanceRecord each in QueryAttendance.GetAttendanceRecords(primaryKeys))
            {
                if (!oneToMany.ContainsKey(each.RefStudentID))
                    oneToMany.Add(each.RefStudentID, new List<AttendanceRecord>());

                oneToMany[each.RefStudentID].Add(each);
            }

            foreach (string each in primaryKeys)
            {
                if (!oneToMany.ContainsKey(each))
                    oneToMany.Add(each, new List<AttendanceRecord>());
            }

            return oneToMany;
        }
    }
}
