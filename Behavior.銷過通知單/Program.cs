using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Presentation;
using FISCA.Permission;
using FISCA.Presentation.Controls;

namespace JHBehavior.銷過通知單
{
    //本功能
    //同時可提供給國高中使用
    //但是需注意
    //權限控管及功能位置需調整開放

    public class Program
    {
        [MainMethod()]
        public static void Main()
        {
            #region 高中
            //RibbonBarItem ClassReports = K12.Presentation.NLDPanels.Class.RibbonBarItems["統計報表"];
            //ClassReports["報表"]["學務相關報表"]["銷過通知單"].Enable = Permissions.班級銷過通知單權限;
            //ClassReports["報表"]["學務相關報表"]["銷過通知單"].Click += delegate
            //{
            //    if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0)
            //    {
            //        new K12.銷過通知單.Report("class").Print();
            //    }
            //    else
            //    {
            //        MsgBox.Show("產生班級報表,請選擇班級!!");
            //    }
            //};

            //RibbonBarItem StudentReports = K12.Presentation.NLDPanels.Student.RibbonBarItems["統計報表"];
            //StudentReports["報表"]["學務相關報表"]["銷過通知單"].Enable = Permissions.學生銷過通知單權限;
            //StudentReports["報表"]["學務相關報表"]["銷過通知單"].Click += delegate
            //{
            //    if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
            //    {
            //        new K12.銷過通知單.Report("student").Print();
            //    }
            //    else
            //    {
            //        MsgBox.Show("產生學生報表,請選擇學生!!");
            //    }
            //};

            ////高中權控
            //Catalog ribbon = RoleAclSource.Instance["學生"]["報表"];
            //ribbon.Add(new RibbonFeature(Permissions.學生銷過通知單, "銷過通知單"));
            //ribbon = RoleAclSource.Instance["班級"]["報表"];
            //ribbon.Add(new RibbonFeature(Permissions.班級銷過通知單, "銷過通知單"));

            #endregion

            #region 國中

            RibbonBarItem JHClassReports = K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"];
            JHClassReports["報表"]["學務相關報表"]["通知單"]["銷過通知單"].Enable = false;
            JHClassReports["報表"]["學務相關報表"]["通知單"]["銷過通知單"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0)
                {
                    new 銷過通知單.Report("class").Print();
                }
                else
                {
                    MsgBox.Show("產生班級報表,請選擇班級!!");
                }
            };

            RibbonBarItem JHStudentReports = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"];
            JHStudentReports["報表"]["學務相關報表"]["通知單"]["銷過通知單"].Enable = false;
            JHStudentReports["報表"]["學務相關報表"]["通知單"]["銷過通知單"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    new 銷過通知單.Report("student").Print();
                }
                else
                {
                    MsgBox.Show("產生學生報表,請選擇學生!!");
                }
            };

            //學生選擇
            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count <= 0)
                {
                    JHStudentReports["報表"]["學務相關報表"]["通知單"]["銷過通知單"].Enable = false;
                }
                else
                {
                    JHStudentReports["報表"]["學務相關報表"]["通知單"]["銷過通知單"].Enable = Permissions.學生銷過通知單權限;
                }
            };

            //班級選擇
            K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count <= 0)
                {
                    JHClassReports["報表"]["學務相關報表"]["通知單"]["銷過通知單"].Enable = false;
                }
                else
                {
                    JHClassReports["報表"]["學務相關報表"]["通知單"]["銷過通知單"].Enable = Permissions.班級銷過通知單權限;
                }
            };



            Framework.Security.Catalog ribbon2 = Framework.Security.RoleAclSource.Instance["學生"]["報表"];
            ribbon2.Add(new Framework.Security.RibbonFeature(Permissions.學生銷過通知單, "銷過通知單"));
            ribbon2 = Framework.Security.RoleAclSource.Instance["班級"]["報表"];
            ribbon2.Add(new Framework.Security.RibbonFeature(Permissions.班級銷過通知單, "銷過通知單"));

            #endregion
        }
    }
}
