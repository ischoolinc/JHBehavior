using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Framework;

namespace JHSchool.Behavior
{
    /// <summary>
    /// 代表目前使用者的相關權限資訊。
    /// </summary>
    public static class Permissions
    {
        /// <summary>
        /// 取得是否有「缺曠」權限。
        /// </summary>
        public static bool Attendance { get { return User.Acl["JHSchool.Student.Ribbon0070"].Executable; } }

        /// <summary>
        /// 取得是否有「長假登錄」權限。
        /// </summary>
        public static bool AttendanceMuti { get { return User.Acl["JHSchool.Student.Ribbon0075"].Executable; } }

        public static string 缺曠資料項目 { get { return "JHSchool.Student.Detail0045"; } }
        public static bool 缺曠資料項目權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[缺曠資料項目].Executable;
            }
        }

        public static string 缺曠學期統計 { get { return "JHSchool.Student.Detail0037"; } }
        public static bool 缺曠學期統計權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[缺曠學期統計].Executable;
            }
        }
    }
}
