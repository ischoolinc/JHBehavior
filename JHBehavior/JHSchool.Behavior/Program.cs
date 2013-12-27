﻿using FISCA;
using FISCA.Presentation;
using Framework;
using Framework.Security;
using JHSchool.Affair;
using JHSchool.Behavior.Legacy;
using JHSchool.Behavior.StuAdminExtendControls;
using JHSchool.Behavior.StuAdminExtendControls.BehaviorStatistics;
using JHSchool.Behavior.StuAdminExtendControls.MeritAndDemeritControl;
using JHSchool.Behavior.StuAdminExtendControls.MoralityMapping;
using JHSchool.Behavior.StuAdminExtendControls.Ribbon;
using JHSchool.Behavior.StudentExtendControls;
using JHSchool.Behavior.StudentExtendControls.Ribbon;
using JHSchool.Data;
using JHSchool.Behavior.ClassExtendControls.Ribbon;

namespace JHSchool.Behavior
{
    public static class Program
    {
        [MainMethod()]
        static public void Main()
        {
            #region 學生的毛毛蟲

            FISCA.Permission.FeatureAce UserPermission;
            //缺曠記錄
            UserPermission = FISCA.Permission.UserAcl.Current[Permissions.缺曠資料項目];
            if (UserPermission.Editable || UserPermission.Viewable)
                Student.Instance.AddDetailBulider(new FISCA.Presentation.DetailBulider<AttendanceItem>());

            //缺曠學期統計
            UserPermission = FISCA.Permission.UserAcl.Current[Permissions.缺曠學期統計];
            if (UserPermission.Editable || UserPermission.Viewable)
                Student.Instance.AddDetailBulider(new FISCA.Presentation.DetailBulider<AttendanceUnifytIItem>());

            #endregion

            #region Attendance

            RibbonBarItem rbItem = Student.Instance.RibbonBarItems["學務"];
            rbItem["缺曠"].Enable = false;
            rbItem["缺曠"].Image = StudentExtendControls.Ribbon.Resources.desk_64;

            rbItem["缺曠"].Click += delegate
            {
                int count = Student.Instance.SelectedList.Count;
                if (count == 1)
                {
                    SingleEditor editor = new SingleEditor(Student.Instance.SelectedList[0]);
                    editor.ShowDialog();
                }
                else if (count > 1)
                {
                    MutiEditor editor = new MutiEditor(K12.Data.Student.SelectByIDs(K12.Presentation.NLDPanels.Student.SelectedSource));
                    editor.ShowDialog();
                }
                else
                {
                    MsgBox.Show("請選擇學生!!");
                }
            };

            if (Permissions.Attendance)
            {
                Student.Instance.ListPaneContexMenu["缺曠"].Enable = false;
                Student.Instance.ListPaneContexMenu["缺曠"].Image = StudentExtendControls.Ribbon.Resources.desk_64;

                Student.Instance.ListPaneContexMenu["缺曠"].Click += delegate
                {
                    if (Student.Instance.SelectedList.Count == 1)
                    {
                        SingleEditor editor = new SingleEditor(Student.Instance.SelectedList[0]);
                        editor.ShowDialog();
                    }
                    else if (Student.Instance.SelectedList.Count > 1)
                    {
                        MutiEditor editor = new MutiEditor(K12.Data.Student.SelectByIDs(K12.Presentation.NLDPanels.Student.SelectedSource));
                        editor.ShowDialog();
                    }
                    else
                    {
                        MsgBox.Show("請選擇學生!!");
                    }
                };

                Student.Instance.SelectedListChanged += delegate
                {
                    if (Permissions.Attendance)
                    {
                        Student.Instance.ListPaneContexMenu["缺曠"].Enable = (Student.Instance.SelectedList.Count >= 1);
                        rbItem["缺曠"].Enable = (Student.Instance.SelectedList.Count >= 1);
                    }
                };
            }

            #endregion

            #region AttendanceMuti

            rbItem = Student.Instance.RibbonBarItems["學務"];
            rbItem["長假登錄"].Enable = false;
            rbItem["長假登錄"].Image = StudentExtendControls.Ribbon.Resources.desk_clock_64;

            rbItem["長假登錄"].Click += delegate
            {
                int count = K12.Presentation.NLDPanels.Student.SelectedSource.Count;
                if (count >= 1)
                {
                    TestSingleEditor SBStatistics = new TestSingleEditor(K12.Presentation.NLDPanels.Student.SelectedSource);
                    SBStatistics.ShowDialog();
                }
                else
                {
                    MsgBox.Show("請選擇學生!!");
                }
            };

            if (Permissions.AttendanceMuti)
            {
                Student.Instance.ListPaneContexMenu["長假登錄"].Enable = false;
                Student.Instance.ListPaneContexMenu["長假登錄"].Image = StudentExtendControls.Ribbon.Resources.desk_clock_64;
                Student.Instance.ListPaneContexMenu["長假登錄"].Click += delegate
                {
                    if (Student.Instance.SelectedList.Count >= 1)
                    {
                        TestSingleEditor SBStatistics = new TestSingleEditor(K12.Presentation.NLDPanels.Student.SelectedSource);
                        SBStatistics.ShowDialog();
                    }
                };

                Student.Instance.SelectedListChanged += delegate
                {
                    if (Permissions.AttendanceMuti)
                    {
                        rbItem["長假登錄"].Enable = (Student.Instance.SelectedList.Count >= 1);
                        Student.Instance.ListPaneContexMenu["長假登錄"].Enable = (Student.Instance.SelectedList.Count >= 1);
                    }
                };
            }

            #endregion

            #region 在班級上增加RibbonBar

            RibbonBarItem classSpecialItem = Class.Instance.RibbonBarItems["學務"];

            classSpecialItem["特殊學生表現"].Size = RibbonBarButton.MenuButtonSize.Medium;
            classSpecialItem["特殊學生表現"].Image = JHSchool.Behavior.ClassExtendControls.Resources.查詢特殊學生;
            classSpecialItem["特殊學生表現"].Enable = (Class.Instance.SelectedList.Count > 0 && User.Acl["JHSchool.Class.Ribbon0060"].Executable);
            classSpecialItem["特殊學生表現"].Click += delegate
            {
                SpecialForm SpecialFormNew = new SpecialForm();
                SpecialFormNew.ShowDialog();
            };

            //classSpecialItem["聯絡資訊管理"].Size = RibbonBarButton.MenuButtonSize.Medium;
            //classSpecialItem["聯絡資訊管理"].Image = Properties.Resources.home_write_64;
            //classSpecialItem["聯絡資訊管理"].Enable = (Class.Instance.SelectedList.Count > 0 && User.Acl["JHBehavior.Class.Ribbon0210"].Executable);
            //classSpecialItem["聯絡資訊管理"].Click += delegate
            //{
            //    AddressEditForm address = new AddressEditForm();
            //    address.ShowDialog();
            //};

            //當選擇班級後,重新檢查一次
            Class.Instance.SelectedListChanged += delegate
            {
                classSpecialItem["特殊學生表現"].Enable = (Class.Instance.SelectedList.Count > 0 && User.Acl["JHSchool.Class.Ribbon0060"].Executable);
                //classSpecialItem["聯絡資訊管理"].Enable = (Class.Instance.SelectedList.Count > 0 && User.Acl["JHBehavior.Class.Ribbon0210"].Executable);
            };

            #endregion

            #region 在學務作業上增加RibbonBar

            RibbonBarItem rbItem1 = StuAdmin.Instance.RibbonBarItems["基本設定"];
            rbItem1.Index = 0;
            rbItem1["管理"].Image = Properties.Resources.network_lock_64;
            rbItem1["管理"].Size = RibbonBarButton.MenuButtonSize.Large;
            rbItem1["管理"]["缺曠類別管理"].Enable = User.Acl["JHSchool.StuAdmin.Ribbon0000"].Executable;
            rbItem1["管理"]["缺曠類別管理"].Click += delegate
            {
                AbsenceConfigForm AbsForm = new AbsenceConfigForm();
                AbsForm.ShowDialog();
            };

            rbItem1["管理"]["每日節次管理"].Enable = User.Acl["JHSchool.StuAdmin.Ribbon0030"].Executable;
            rbItem1["管理"]["每日節次管理"].Click += delegate
            {
                PeriodConfigForm PerForm = new PeriodConfigForm();
                PerForm.ShowDialog();
            };

            rbItem1["管理"]["功過換算管理"].Enable = User.Acl["JHSchool.StuAdmin.Ribbon0010"].Executable;
            rbItem1["管理"]["功過換算管理"].Click += delegate
            {
                ReduceForm RedForm = new ReduceForm();
                RedForm.ShowDialog();
            };

            rbItem1["對照/代碼"].Image = Properties.Resources.notepad_lock_64;
            rbItem1["對照/代碼"].Size = RibbonBarButton.MenuButtonSize.Large;
            rbItem1["對照/代碼"]["導師評語代碼表"].Enable = User.Acl["JHSchool.StuAdmin.Ribbon0090"].Executable;
            rbItem1["對照/代碼"]["導師評語代碼表"].Click += delegate
            {
                MoralityForm DiscForm = new MoralityForm();
                DiscForm.ShowDialog();
            };

            rbItem1["對照/代碼"]["表現程度代碼表"].Enable = User.Acl["JHSchool.StuAdmin.Ribbon0022"].Executable;
            rbItem1["對照/代碼"]["表現程度代碼表"].Click += delegate
            {
                PerformanceDegreeForm foor = new PerformanceDegreeForm();
                foor.ShowDialog();
            };

            rbItem1["對照/代碼"]["獎懲事由代碼表"].Enable = User.Acl["JHSchool.StuAdmin.Ribbon0040"].Executable;
            rbItem1["對照/代碼"]["獎懲事由代碼表"].Click += delegate
            {
                DisciplineForm DiscForm = new DisciplineForm();
                DiscForm.ShowDialog();
            };

            //RibbonBarItem rbItem3 = StuAdmin.Instance.RibbonBarItems["設定"];

            rbItem1["設定"].Image = Properties.Resources.sandglass_unlock_64;
            rbItem1["設定"].Size = RibbonBarButton.MenuButtonSize.Large;
            rbItem1["設定"]["日常生活表現輸入時間設定"].Enable = User.Acl["JHSchool.StuAdmin.Ribbon0025"].Executable;
            rbItem1["設定"]["日常生活表現輸入時間設定"].Click += delegate
            {
                DailyLifeInputControl roto = new DailyLifeInputControl();
                roto.ShowDialog();
            };

            rbItem1["設定"]["上課天數設定"].Enable = User.Acl["JJHSchool.StuAdmin.Ribbon0021"].Executable;
            rbItem1["設定"]["上課天數設定"].Click += delegate
            {
                new SchoolDayForm().ShowDialog();
            };

            RibbonBarItem batchEdit = StuAdmin.Instance.RibbonBarItems["批次作業/查詢"];
            batchEdit.Index = 1;
            batchEdit["獎懲批次修改"].Image = Properties.Resources.star_write_64;
            batchEdit["獎懲批次修改"].Enable = User.Acl["JHSchool.StuAdmin.Ribbon0050"].Executable;
            batchEdit["獎懲批次修改"].Click += delegate
            {
                MerDemEditForm DiscipTotal = new MerDemEditForm();
                DiscipTotal.ShowDialog();
            };

            //batchEdit["缺曠資料檢視"].Image = Properties.Resources.preview_64;
            //batchEdit["缺曠資料檢視"].Enable = User.Acl["JHSchool.StuAdmin.Ribbon0060"].Executable;
            //batchEdit["缺曠資料檢視"].Click += delegate
            //{
            //    AttendanceEditForm AttendanceTotal = new AttendanceEditForm();
            //    AttendanceTotal.ShowDialog();
            //};

            //RibbonBarItem batchEdit1 = StuAdmin.Instance.RibbonBarItems["資料統計"];
            //batchEdit1["報表"].Image = Properties.Resources.paste_64;
            //batchEdit1["報表"]["銷過記錄清單"].Enable = User.Acl["JHSchool.StuAdmin.Ribbon0100"].Executable;
            //batchEdit1["報表"]["銷過記錄清單"].Click += delegate
            //{
            //    StudentDemeritClear clearData = new StudentDemeritClear();
            //    clearData.ShowDialog();
            //};

            #endregion

            #region 在教務作業上增加RibbonBar






            //EduAdmin 教務作業 






            #endregion

            #region 匯出/匯入

            //匯出
            RibbonBarButton rbItemExport = Student.Instance.RibbonBarItems["資料統計"]["匯出"];
            rbItemExport["學務相關匯出"]["匯出缺曠記錄"].Enable = User.Acl["JHSchool.Student.Ribbon0150"].Executable;
            rbItemExport["學務相關匯出"]["匯出缺曠記錄"].Click += delegate
            {
                SmartSchool.API.PlugIn.Export.Exporter exporter = new JHSchool.Behavior.ImportExport.ExportAttendance();
                JHSchool.Behavior.ImportExport.ExportStudentV2 wizard = new JHSchool.Behavior.ImportExport.ExportStudentV2(exporter.Text, exporter.Image);
                exporter.InitializeExport(wizard);
                wizard.ShowDialog();
            };

            rbItemExport["學務相關匯出"]["匯出獎勵記錄"].Enable = User.Acl["JHSchool.Student.Ribbon0152"].Executable;
            rbItemExport["學務相關匯出"]["匯出獎勵記錄"].Click += delegate
            {
                SmartSchool.API.PlugIn.Export.Exporter exporter = new JHSchool.Behavior.ImportExport.ExportMerit();
                JHSchool.Behavior.ImportExport.ExportStudentV2 wizard = new JHSchool.Behavior.ImportExport.ExportStudentV2(exporter.Text, exporter.Image);
                exporter.InitializeExport(wizard);
                wizard.ShowDialog();
            };

            rbItemExport["學務相關匯出"]["匯出懲戒記錄"].Enable = User.Acl["JHSchool.Student.Ribbon0154"].Executable;
            rbItemExport["學務相關匯出"]["匯出懲戒記錄"].Click += delegate
            {
                SmartSchool.API.PlugIn.Export.Exporter exporter = new JHSchool.Behavior.ImportExport.ExportDemerit();
                JHSchool.Behavior.ImportExport.ExportStudentV2 wizard = new JHSchool.Behavior.ImportExport.ExportStudentV2(exporter.Text, exporter.Image);
                exporter.InitializeExport(wizard);
                wizard.ShowDialog();
            };

            rbItemExport["學務相關匯出"]["匯出獎勵懲戒記錄"].Enable = User.Acl["JHSchool.Student.Ribbon0156"].Executable;
            rbItemExport["學務相關匯出"]["匯出獎勵懲戒記錄"].Click += delegate
            {
                SmartSchool.API.PlugIn.Export.Exporter exporter = new JHSchool.Behavior.ImportExport.ExportDiscipline();
                JHSchool.Behavior.ImportExport.ExportStudentV2 wizard = new JHSchool.Behavior.ImportExport.ExportStudentV2(exporter.Text, exporter.Image);
                exporter.InitializeExport(wizard);
                wizard.ShowDialog();
            };

            //暫時註解
            rbItemExport["學務相關匯出"]["匯出缺曠統計"].Enable = User.Acl["JHSchool.Student.Ribbon0158"].Executable;
            rbItemExport["學務相關匯出"]["匯出缺曠統計"].Click += delegate
            {
                SmartSchool.API.PlugIn.Export.Exporter exporter = new JHSchool.Behavior.ImportExport.ExportAttendanceStatistics();
                JHSchool.Behavior.ImportExport.ExportStudentV2 wizard = new JHSchool.Behavior.ImportExport.ExportStudentV2(exporter.Text, exporter.Image);
                exporter.InitializeExport(wizard);
                wizard.ShowDialog();
            };

            //暫時註解
            rbItemExport["學務相關匯出"]["匯出獎勵懲戒統計"].Enable = User.Acl["JHSchool.Student.Ribbon0160"].Executable;
            rbItemExport["學務相關匯出"]["匯出獎勵懲戒統計"].Click += delegate
            {
                SmartSchool.API.PlugIn.Export.Exporter exporter = new JHSchool.Behavior.ImportExport.ExportDisciplineStatistics();
                JHSchool.Behavior.ImportExport.ExportStudentV2 wizard = new JHSchool.Behavior.ImportExport.ExportStudentV2(exporter.Text, exporter.Image);
                exporter.InitializeExport(wizard);
                wizard.ShowDialog();
            };

            //匯入
            RibbonBarButton rbItemImport = Student.Instance.RibbonBarItems["資料統計"]["匯入"];
            rbItemImport["學務相關匯入"]["匯入缺曠記錄"].Enable = User.Acl["JHSchool.Student.Ribbon0151"].Executable;
            rbItemImport["學務相關匯入"]["匯入缺曠記錄"].Click += delegate
            {
                SmartSchool.API.PlugIn.Import.Importer importer = new JHSchool.Behavior.ImportExport.ImportAttendance();
                JHSchool.Behavior.ImportExport.ImportStudentV2 wizard = new JHSchool.Behavior.ImportExport.ImportStudentV2(importer.Text, importer.Image);
                importer.InitializeImport(wizard);
                wizard.ShowDialog();
            };

            rbItemImport["學務相關匯入"]["匯入獎勵懲戒記錄"].Enable = User.Acl["JHSchool.Student.Ribbon0157"].Executable;
            rbItemImport["學務相關匯入"]["匯入獎勵懲戒記錄"].Click += delegate
            {
                SmartSchool.API.PlugIn.Import.Importer importer = new JHSchool.Behavior.ImportExport.ImportDiscipline();
                JHSchool.Behavior.ImportExport.ImportStudentV2 wizard = new JHSchool.Behavior.ImportExport.ImportStudentV2(importer.Text, importer.Image);
                importer.InitializeImport(wizard);
                wizard.ShowDialog();
            };

            rbItemImport["學務相關匯入"]["匯入缺曠統計"].Enable = User.Acl["JHSchool.Student.Ribbon0159"].Executable;
            rbItemImport["學務相關匯入"]["匯入缺曠統計"].Click += delegate
            {
                SmartSchool.API.PlugIn.Import.Importer importer = new JHSchool.Behavior.ImportExport.ImportAttendanceStatistics();
                JHSchool.Behavior.ImportExport.ImportStudentV2 wizard = new JHSchool.Behavior.ImportExport.ImportStudentV2(importer.Text, importer.Image);
                importer.InitializeImport(wizard);
                wizard.ShowDialog();
            };

            rbItemImport["學務相關匯入"]["匯入獎勵懲戒統計"].Enable = User.Acl["JHSchool.Student.Ribbon0161"].Executable;
            rbItemImport["學務相關匯入"]["匯入獎勵懲戒統計"].Click += delegate
            {
                SmartSchool.API.PlugIn.Import.Importer importer = new JHSchool.Behavior.ImportExport.ImportDisciplineStatistics();
                JHSchool.Behavior.ImportExport.ImportStudentV2 wizard = new JHSchool.Behavior.ImportExport.ImportStudentV2(importer.Text, importer.Image);
                importer.InitializeImport(wizard);
                wizard.ShowDialog();
            };
            #endregion

            #region 註冊權限管理
            Catalog ribbon = RoleAclSource.Instance["學生"]["功能按鈕"];
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0070", "缺曠"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0075", "長假登錄"));

            //暫時註解
            //ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0085", "缺曠獎懲資料統計"));

            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0150", "匯出缺曠記錄"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0152", "匯出獎勵記錄"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0154", "匯出懲戒記錄"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0156", "匯出獎勵懲戒記錄"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0158", "匯出缺曠統計"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0160", "匯出獎勵懲戒統計"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0151", "匯入缺曠記錄"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0157", "匯入獎勵懲戒記錄"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0159", "匯入缺曠統計"));
            ribbon.Add(new RibbonFeature("JHSchool.Student.Ribbon0161", "匯入獎勵懲戒統計"));

            Catalog detail = RoleAclSource.Instance["學生"]["資料項目"];
            detail.Add(new DetailItemFeature(typeof(AttendanceItem)));
            detail.Add(new DetailItemFeature(typeof(AttendanceUnifytIItem))); //缺曠學期統計(NEW)

            //暫時註解
            //detail.Add(new DetailItemFeature(typeof(StudentExtendControls.AttendanceStatisticsItem)));
            //detail.Add(new DetailItemFeature(typeof(StudentExtendControls.MDStatisticsItem)));

            //detail.Add(new DetailItemFeature(typeof(StudentExtendControls.MeritUnifytItem))); //獎勵整合
            //detail.Add(new DetailItemFeature(typeof(StudentExtendControls.DemeritUnifytItem))); //懲戒整合
            //detail.Add(new DetailItemFeature(typeof(StudentExtendControls.AttendanceUnifyItem))); //缺曠整合

            ribbon = RoleAclSource.Instance["班級"]["功能按鈕"];
            ribbon.Add(new RibbonFeature("JHSchool.Class.Ribbon0060", "特殊學生表現"));
            //ribbon.Add(new RibbonFeature("JHBehavior.Class.Ribbon0210", "聯絡資訊管理"));

            //學務作業
            ribbon = RoleAclSource.Instance["學務作業"];
            ribbon.Add(new RibbonFeature("JHSchool.StuAdmin.Ribbon0000", "缺曠類別管理"));
            ribbon.Add(new RibbonFeature("JHSchool.StuAdmin.Ribbon0010", "功過換算管理"));
            ribbon.Add(new RibbonFeature("JHSchool.StuAdmin.Ribbon0030", "每日節次管理"));
            ribbon.Add(new RibbonFeature("JHSchool.StuAdmin.Ribbon0040", "獎懲事由管理"));

            ribbon.Add(new RibbonFeature("JHSchool.StuAdmin.Ribbon0050", "獎懲批次修改"));
            //ribbon.Add(new RibbonFeature("JHSchool.StuAdmin.Ribbon0060", "缺曠資料檢視"));

            ribbon.Add(new RibbonFeature("JHSchool.StuAdmin.Ribbon0090", "導師評語代碼表"));

            //ribbon.Add(new RibbonFeature("JHSchool.StuAdmin.Ribbon0100", "銷過記錄清單"));

            //暫時註解
            //ribbon.Add(new RibbonFeature("JHSchool.StuAdmin.Ribbon0100", "缺曠獎懲資料統計"));

            ribbon.Add(new RibbonFeature("JHSchool.StuAdmin.Ribbon0022", "表現程度對照表"));
            ribbon.Add(new RibbonFeature("JHSchool.StuAdmin.Ribbon0025", "日常生活表現輸入時間設定"));

            Catalog toolSetup = RoleAclSource.Instance["學務作業"];
            toolSetup.Add(new RibbonFeature("JJHSchool.StuAdmin.Ribbon0021", "上課天數設定"));
            #endregion
        }
    }
}