using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Presentation;
using Framework.Security;
using Framework;

namespace Behavior.MeritDemeritStatistics
{
    public static class Program
    {
        [MainMethod("JHBehavior.MeritDemeritStatistics")]
        static public void Main()
        {
            RibbonBarItem rbItem = MotherForm.RibbonBarItems["學務作業", "資料統計"];

            rbItem["報表"]["獎懲人數統計"].Enable = User.Acl["JHBehavior.MeritDemeritStatistics"].Executable;
            //rbItem["報表"]["獎懲人數統計"].Image = Properties.Resources.attach_write_64;
            rbItem["報表"]["獎懲人數統計"].Click += delegate
            {
                SchoolYearAndSemesterSelectForm SYASSF = new SchoolYearAndSemesterSelectForm();
                SYASSF.ShowDialog();
            };

            Catalog ribbon = RoleAclSource.Instance["學務作業"];
            ribbon.Add(new RibbonFeature("JHBehavior.MeritDemeritStatistics", "獎懲人數統計"));

        }
    }
}
