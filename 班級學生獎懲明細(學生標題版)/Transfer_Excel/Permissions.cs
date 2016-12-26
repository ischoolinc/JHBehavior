using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Transfer_Excel
{
    class Permissions
    {
        public static string 班級學生獎懲明細 { get { return "Transfer_Excel.TitleDiscipline"; } }

        public static bool 班級學生獎懲明細權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[班級學生獎懲明細].Executable;
            }
        }
    }
}
