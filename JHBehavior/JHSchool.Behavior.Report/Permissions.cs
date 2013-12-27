using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Framework;

namespace JHSchool.Behavior.Report
{
    /// <summary>
    /// 代表目前使用者的相關權限資訊。
    /// </summary>
    public static class Permissions
    {
        //毛毛蟲權限(對應於毛毛蟲內之權限字串)
        public static string 學生缺曠通知單 { get { return "JHSchool.Student.Report0020"; } }
        public static string 學生獎懲通知單 { get { return "JHSchool.Student.Report0030"; } }
        public static string 學生獎勵懲戒明細 { get { return "JHSchool.Student.Report0040"; } }
        public static string 學生獎勵明細 { get { return "JHSchool.Student.Report0040.5"; } }
        public static string 學生缺曠明細 { get { return "JHSchool.Student.Report0050"; } }

        public static string 學生懲戒通知單 { get { return "JHSchool.Student.Report0045.20130227"; } }
        public static bool 學生懲戒通知單權限
        {
            get { return User.Acl[學生懲戒通知單].Executable; }
        }

        public static bool 學生缺曠通知單權限
        {
            get { return User.Acl[學生缺曠通知單].Executable; }
        }

        public static bool 學生獎懲通知單權限
        {
            get { return User.Acl[學生獎懲通知單].Executable; }
        }

        public static bool 學生獎勵懲戒明細權限
        {
            get { return User.Acl[學生獎勵懲戒明細].Executable; }
        }

        public static bool 學生獎勵明細權限
        {
            get { return User.Acl[學生獎勵明細].Executable; }
        }

        public static bool 學生缺曠明細權限
        {
            get { return User.Acl[學生缺曠明細].Executable; }
        }


        public static string 班級點名表 { get { return "JHSchool.Class.Report0020"; } }
        public static string 班級點名表_自定樣板 { get { return "JHSchool.Class.Report0020.1"; } }
        public static string 班級通訊錄 { get { return "JHSchool.Class.Report0030"; } }

        public static bool 班級點名表權限
        {
            get { return User.Acl[班級點名表].Executable; }
        }

        public static bool 班級點名表_自定樣板_權限
        {
            get { return User.Acl[班級點名表_自定樣板].Executable; }
        }

        public static bool 班級通訊錄權限
        {
            get { return User.Acl[班級通訊錄].Executable; }
        }

        public static string 班級缺曠通知單 { get { return "JHSchool.Class.Report0070"; } }
        public static string 班級懲戒通知單 { get { return "JHSchool.Class.Report0045"; } }
        public static string 獎勵懲戒通知單 { get { return "JHSchool.Class.Report0040"; } }
        public static string 班級學生缺曠統計 { get { return "JHSchool.Class.Report0090"; } }
        public static string 班級學生獎懲統計 { get { return "JHSchool.Class.Report0095"; } }

        public static bool 班級缺曠通知單權限
        {
            get { return User.Acl[班級缺曠通知單].Executable; }
        }

        public static bool 班級懲戒通知單權限
        {
            get { return User.Acl[班級懲戒通知單].Executable; }
        }

        public static bool 獎勵懲戒通知單權限
        {
            get { return User.Acl[獎勵懲戒通知單].Executable; }
        }

        public static bool 班級學生缺曠統計權限
        {
            get { return User.Acl[班級學生缺曠統計].Executable; }
        }

        public static bool 班級學生獎懲統計權限
        {
            get { return User.Acl[班級學生獎懲統計].Executable; }
        }



        public static string 班級獎懲記錄明細 { get { return "JHSchool.Class.Report0050"; } }
        public static string 班級缺曠記錄明細 { get { return "JHSchool.Class.Report0080"; } }

        public static bool 班級獎懲記錄明細權限
        {
            get { return User.Acl[班級獎懲記錄明細].Executable; }
        }

        public static bool 班級缺曠記錄明細權限
        {
            get { return User.Acl[班級缺曠記錄明細].Executable; }
        }

        public static string 獎懲週報表 { get { return "JHSchool.Class.Report0060"; } }
        public static string 缺曠週報表_依節次 { get { return "JHSchool.Class.Report0100"; } }
        public static string 缺曠週報表_依假別 { get { return "JHSchool.Class.Report0110"; } }

        public static bool 獎懲週報表權限
        {
            get { return User.Acl[獎懲週報表].Executable; }
        }

        public static bool 缺曠週報表_依節次權限
        {
            get { return User.Acl[缺曠週報表_依節次].Executable; }
        }

        public static bool 缺曠週報表_依假別權限
        {
            get { return User.Acl[缺曠週報表_依假別].Executable; }
        }






    }
}
