using FISCA;
using FISCA.Permission;
using FISCA.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Transfer_Excel
{
    public class Program
    {
        [MainMethod]
        public static void Main()
        {
            RibbonBarItem classSpecialItem = K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"];
            classSpecialItem["報表"]["學務相關報表"]["班級學生獎懲明細(學生標題版)"].Enable = Permissions.班級學生獎懲明細權限;
            classSpecialItem["報表"]["學務相關報表"]["班級學生獎懲明細(學生標題版)"].Click += delegate
            {
                new Transfer_Excel.Report().Print();
            };

            Catalog ribbon = RoleAclSource.Instance["班級"]["報表"];
            ribbon.Add(new RibbonFeature(Permissions.班級學生獎懲明細, "班級學生獎懲明細(學生標題版)"));
        }
    }
}
