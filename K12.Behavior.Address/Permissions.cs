using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Behavior.Address
{
    /// <summary>
    /// 代表目前使用者的相關權限資訊。
    /// </summary>
    public static class Permissions
    {
        public static string 聯絡資訊管理 { get { return "JHBehavior.Class.Ribbon0210"; } }

        public static bool 聯絡資訊管理權限
        {
            get { return FISCA.Permission.UserAcl.Current[聯絡資訊管理].Executable; }
        }


    }
}
