using FISCA;
using Framework;
using Framework.Security;
using JHSchool.Data;
using System.Collections.Generic;

namespace JHSchool.Behavior.Report
{
    public static class Program
    {
        [MainMethod()]
        public static void Main()
        {
            ClassFalse();
            StudentFalse();

            Class.Instance.SelectedListChanged += delegate
            {
                if (Class.Instance.SelectedKeys.Count <= 0)
                {
                    ClassFalse();
                }
                else
                {
                    Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級點名表"].Enable = Permissions.班級點名表權限 && Class.Instance.SelectedKeys.Count > 0;
                    Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級點名表(自定樣板)"].Enable = Permissions.班級點名表_自定樣板_權限 && Class.Instance.SelectedKeys.Count > 0;
                    //Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級通訊錄"].Enable = Permissions.班級通訊錄權限 && Class.Instance.SelectedKeys.Count > 0;
                    Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["通知單"]["缺曠通知單"].Enable = Permissions.班級缺曠通知單權限 && Class.Instance.SelectedKeys.Count > 0;
                    Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["通知單"]["懲戒通知單"].Enable = Permissions.班級懲戒通知單權限 && Class.Instance.SelectedKeys.Count > 0;
                    Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["通知單"]["獎懲通知單"].Enable = Permissions.獎勵懲戒通知單權限 && Class.Instance.SelectedKeys.Count > 0;
                    Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級學生缺曠統計"].Enable = Permissions.班級學生缺曠統計權限 && Class.Instance.SelectedKeys.Count > 0;
                    Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級學生獎懲統計"].Enable = Permissions.班級學生獎懲統計權限 && Class.Instance.SelectedKeys.Count > 0;
                    Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級獎懲記錄明細"].Enable = Permissions.班級獎懲記錄明細權限 && Class.Instance.SelectedKeys.Count > 0;
                    Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級缺曠記錄明細"].Enable = Permissions.班級缺曠記錄明細權限 && Class.Instance.SelectedKeys.Count > 0;
                    Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["獎懲週報表"].Enable = Permissions.獎懲週報表權限 && Class.Instance.SelectedKeys.Count > 0;
                    Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["缺曠週報表(依節次)"].Enable = Permissions.缺曠週報表_依節次權限 && Class.Instance.SelectedKeys.Count > 0;
                    Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["缺曠週報表(依假別)"].Enable = Permissions.缺曠週報表_依假別權限 && Class.Instance.SelectedKeys.Count > 0;
                }
            };

            #region 學生狀態
            Student.Instance.SelectedListChanged += delegate
            {
                if (Student.Instance.SelectedKeys.Count <= 0)
                {
                    StudentFalse();
                }
                else
                {
                    //20130227 - new
                    Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["通知單"]["懲戒通知單"].Enable = Permissions.學生懲戒通知單權限 && Student.Instance.SelectedKeys.Count > 0;
                    Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["通知單"]["缺曠通知單"].Enable = Permissions.學生缺曠通知單權限 && Student.Instance.SelectedKeys.Count > 0;
                    Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["通知單"]["獎懲通知單"].Enable = Permissions.學生獎懲通知單權限 && Student.Instance.SelectedKeys.Count > 0;
                    Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["學生獎勵懲戒明細"].Enable = Permissions.學生獎勵懲戒明細權限 && Student.Instance.SelectedKeys.Count > 0;
                    Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["學生獎勵明細"].Enable = Permissions.學生獎勵明細權限 && Student.Instance.SelectedKeys.Count > 0;
                    Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["學生缺曠明細"].Enable = Permissions.學生缺曠明細權限 && Student.Instance.SelectedKeys.Count > 0;
                }
            };
            #endregion

            #region 班級
            //Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級點名表"].Enable = Permissions.班級點名表權限;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級點名表"].Click += delegate
            {
                new 班級點名表.Report().Print();
            };

            //Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級點名表(自定樣板)"].Enable = Permissions.班級點名表_自定樣板_權限;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級點名表(自定樣板)"].Click += delegate
            {
                SpecialRollCallForm SRCForm = new SpecialRollCallForm(K12.Presentation.NLDPanels.Class.SelectedSource);
                SRCForm.ShowDialog();
            };

            //Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級通訊錄"].Enable = Permissions.班級通訊錄權限;
            //Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級通訊錄"].Click += delegate
            //{
            //    new 班級通訊錄.Report().Print();
            //};

            //Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["缺曠通知單"].Enable = Permissions.班級缺曠通知單權限;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["通知單"]["缺曠通知單"].Click += delegate
            {
                new 缺曠通知單.Report("class").Print();
            };

            //Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["懲戒通知單"].Enable = Permissions.班級懲戒通知單權限;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["通知單"]["懲戒通知單"].Click += delegate
            {
                new 懲戒通知單.Report("class").Print();
            };

            //Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["獎勵懲戒通知單"].Enable = Permissions.獎勵懲戒通知單權限;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["通知單"]["獎懲通知單"].Click += delegate
            {
                new 獎勵懲戒通知單.Report("class").Print();
            };

            //new
            //Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級學生缺曠統計"].Enable = Permissions.班級學生缺曠統計權限;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級學生缺曠統計"].Click += delegate
            {
                new 班級學生缺曠統計.Report().Print();
            };

            //new
            //Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級學生獎懲統計"].Enable = Permissions.班級學生獎懲統計權限;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級學生獎懲統計"].Click += delegate
            {
                new 班級學生獎懲統計.Report().Print();
            };

            //Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級獎懲記錄明細"].Enable = Permissions.班級獎懲記錄明細權限;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級獎懲記錄明細"].Click += delegate
            {
                new 班級獎懲記錄明細.Report().Print();
            };

            //Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級缺曠記錄明細"].Enable = Permissions.班級缺曠記錄明細權限;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級缺曠記錄明細"].Click += delegate
            {
                new 班級學生缺曠明細.Report().Print();
            };

            //Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["獎懲週報表"].Enable = Permissions.獎懲週報表權限;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["獎懲週報表"].Click += delegate
            {
                new 獎懲週報表.Report().Print();
            };

            //Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["缺曠週報表(依節次)"].Enable = Permissions.缺曠週報表_依節次權限;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["缺曠週報表(依節次)"].Click += delegate
            {
                new 缺曠週報表_依節次.Report().Print();                
            };

            //Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["缺曠週報表(依假別)"].Enable = Permissions.缺曠週報表_依假別權限;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["缺曠週報表(依假別)"].Click += delegate
            {
                new 缺曠週報表_依假別.Report().Print();
            };
            #endregion

            #region 學生
            //Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["缺曠通知單"].Enable = Permissions.學生缺曠通知單權限;
            Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["通知單"]["缺曠通知單"].Click += delegate
            {
                new 缺曠通知單.Report("student").Print();
            };

            Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["通知單"]["懲戒通知單"].Click += delegate
            {
                new 懲戒通知單.Report("student").Print();
            };

            //Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["點名表(Word)"].Enable = User.Acl["JHSchool.Class.Report0020.2"].Executable;
            //Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["點名表(Word)"].Click += delegate
            //{
            //    SpecialRollCallForm SRCForm = new SpecialRollCallForm(JHStudent.SelectByIDs(K12.Presentation.NLDPanels.Student.SelectedSource));
            //    SRCForm.ShowDialog();
            //};

            //Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["獎懲通知單"].Enable = Permissions.學生獎懲通知單權限;
            Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["通知單"]["獎懲通知單"].Click += delegate
            {
                new 獎勵懲戒通知單.Report("student").Print();
            };

            //Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["學生獎勵懲戒明細"].Enable = Permissions.學生獎懲明細權限;
            Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["學生獎勵懲戒明細"].Click += delegate
            {
                new 學生獎懲明細.Report().Print();
            };

            Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["學生獎勵明細"].Click += delegate
            {
                new 學生獎勵明細.Report().Print();
            };

            //Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["學生缺曠明細"].Enable = Permissions.學生缺曠明細權限;
            Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["學生缺曠明細"].Click += delegate
            {
                new 學生缺曠明細.Report().Print();
            };
            #endregion

            #region 註冊權限
            RoleAclSource.Instance["學生"]["報表"].Add(new ReportFeature(Permissions.學生缺曠通知單, "缺曠通知單"));
            //20130227 - new
            RoleAclSource.Instance["學生"]["報表"].Add(new ReportFeature(Permissions.學生懲戒通知單, "懲戒通知單"));
            RoleAclSource.Instance["學生"]["報表"].Add(new ReportFeature(Permissions.學生獎懲通知單, "獎懲通知單"));
            RoleAclSource.Instance["學生"]["報表"].Add(new ReportFeature(Permissions.學生獎勵懲戒明細, "學生獎勵懲戒明細"));
            RoleAclSource.Instance["學生"]["報表"].Add(new ReportFeature(Permissions.學生獎勵明細, "學生獎勵明細"));
            RoleAclSource.Instance["學生"]["報表"].Add(new ReportFeature(Permissions.學生缺曠明細, "學生缺曠明細"));

            RoleAclSource.Instance["班級"]["報表"].Add(new ReportFeature(Permissions.班級點名表, "班級點名表"));
            RoleAclSource.Instance["班級"]["報表"].Add(new ReportFeature(Permissions.班級點名表_自定樣板, "班級點名表(自定樣板)"));
            //RoleAclSource.Instance["學生"]["報表"].Add(new ReportFeature("JHSchool.Class.Report0020.2", "點名表(Word)"));
            //RoleAclSource.Instance["班級"]["報表"].Add(new ReportFeature(Permissions.班級通訊錄, "班級通訊錄"));

            RoleAclSource.Instance["班級"]["報表"].Add(new ReportFeature(Permissions.班級缺曠通知單, "缺曠通知單"));
            RoleAclSource.Instance["班級"]["報表"].Add(new ReportFeature(Permissions.班級懲戒通知單, "懲戒通知單"));
            RoleAclSource.Instance["班級"]["報表"].Add(new ReportFeature(Permissions.獎勵懲戒通知單, "獎勵懲戒通知單"));

            //星期一繼續處理權限控管
            RoleAclSource.Instance["班級"]["報表"].Add(new ReportFeature(Permissions.班級學生缺曠統計, "班級學生缺曠統計"));
            RoleAclSource.Instance["班級"]["報表"].Add(new ReportFeature(Permissions.班級學生獎懲統計, "班級學生獎懲統計"));

            RoleAclSource.Instance["班級"]["報表"].Add(new ReportFeature(Permissions.班級獎懲記錄明細, "班級獎懲記錄明細"));
            RoleAclSource.Instance["班級"]["報表"].Add(new ReportFeature(Permissions.班級缺曠記錄明細, "班級缺曠記錄明細"));

            RoleAclSource.Instance["班級"]["報表"].Add(new ReportFeature(Permissions.獎懲週報表, "獎懲週報表"));
            RoleAclSource.Instance["班級"]["報表"].Add(new ReportFeature(Permissions.缺曠週報表_依節次, "缺曠週報表(依節次)"));
            RoleAclSource.Instance["班級"]["報表"].Add(new ReportFeature(Permissions.缺曠週報表_依假別, "缺曠週報表(依假別)"));
            #endregion
        }

        private static void ClassFalse()
        {
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級點名表"].Enable = false;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級點名表(自定樣板)"].Enable = false;
            //Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級通訊錄"].Enable = false;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["通知單"]["缺曠通知單"].Enable = false;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["通知單"]["懲戒通知單"].Enable = false;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["通知單"]["獎懲通知單"].Enable = false;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級學生缺曠統計"].Enable = false;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級學生獎懲統計"].Enable = false;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級獎懲記錄明細"].Enable = false;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["班級缺曠記錄明細"].Enable = false;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["獎懲週報表"].Enable = false;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["缺曠週報表(依節次)"].Enable = false;
            Class.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["缺曠週報表(依假別)"].Enable = false;
        }

        private static void StudentFalse()
        {
            Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["通知單"]["懲戒通知單"].Enable = false;
            Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["通知單"]["缺曠通知單"].Enable = false;
            Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["通知單"]["獎懲通知單"].Enable = false;
            Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["學生獎勵懲戒明細"].Enable = false;
            Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["學生獎勵明細"].Enable = false;
            Student.Instance.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["學生缺曠明細"].Enable = false;
        }
    }
}
