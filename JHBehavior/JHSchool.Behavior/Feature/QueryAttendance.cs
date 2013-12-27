using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using Framework;
using FISCA.DSAUtil;
using FISCA.Authentication;

namespace JHSchool.Behavior.Feature
{
    /// <summary>
    /// 取得學生缺曠紀錄
    /// </summary>
    [AutoRetryOnWebException()]
    public static class QueryAttendance
    {
        private const string GET_ATTENDANCE = "SmartSchool.Student.Attendance.GetAttendance";

        /// <summary>
        /// 取得所有的缺曠紀錄
        /// </summary>
        /// <returns></returns>
        public static List<AttendanceRecord> GetAllAttendanceRecord()
        {
            StringBuilder req = new StringBuilder("<Request><Field><All/></Field></Request><Order><SchoolYear/><Semester/><OccurDate/></Order>");
            List<AttendanceRecord> result = new List<AttendanceRecord>();

            foreach (XmlElement item in DSAServices.CallService(GET_ATTENDANCE, new DSRequest(req.ToString())).GetContent().GetElements("Attendance"))
            {
                result.Add(new AttendanceRecord(item));
            }
            
            return result;
        }

        /// <summary>
        /// 取得指定的學生的缺曠紀錄
        /// </summary>
        /// <param name="primaryKeys">學生ID的集合</param>
        /// <returns></returns>
        public static List<AttendanceRecord> GetAttendanceRecords(IEnumerable<string> primaryKeys)
        {
            bool haskey = false;

            StringBuilder req = new StringBuilder("<Request><Field><All/></Field><Condition>");
            
            foreach (string key in primaryKeys)
            {
                req.Append("<RefStudentID>" + key + "</RefStudentID>");
                haskey = true;
            }

            req.Append("</Condition><Order><SchoolYear/><Semester/><OccurDate/></Order></Request>");

            List<AttendanceRecord> result = new List<AttendanceRecord>();
            
            if (haskey)
            {
                foreach (XmlElement item in DSAServices.CallService(GET_ATTENDANCE, new DSRequest(req.ToString())).GetContent().GetElements("Attendance"))
                {
                    result.Add(new AttendanceRecord(item));
                }
            }

            return result;
        }    


        public static DSResponse GetAttendance(DSRequest request)
        {
            return DSAServices.CallService("SmartSchool.Student.Attendance.GetAttendance", request);
        }

        /// <summary>
        /// 取得缺曠累計,超過設定值的學生名單
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        public static DSResponse GetAttendanceStatistic(DSRequest request)
        {
            return DSAServices.CallService("SmartSchool.Student.Attendance.GetAbsenceStatistic", request);
        }

        /// <summary>
        /// 取得全勤的學生名單
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        public static DSResponse GetNoAbsenceStatistic(DSRequest request)
        {
            return DSAServices.CallService("SmartSchool.Student.Attendance.GetNoAbsenceStatistic", request);
        }
    }
}
