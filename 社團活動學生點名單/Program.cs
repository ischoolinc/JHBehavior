using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Presentation;
using JHSchool;
using Framework;
using Framework.Security;

namespace Behavior.ClubActivitiesPointList
{
    public class Program
    {
        [MainMethod()]
        public static void Main()
        {

            #region 課程功能

            //設定起始
            Course.Instance.RibbonBarItems["資料統計"]["報表"]["社團相關報表"]["社團活動學生點名單"].Enable = (Course.Instance.SelectedList.Count > 0 && User.Acl["JHBehavior.Course.Ribbon0210"].Executable);
            //設定當執行後
            Course.Instance.RibbonBarItems["資料統計"]["報表"]["社團相關報表"]["社團活動學生點名單"].Click += delegate
            {
                ClubActivitiesForm ClubPointList = new ClubActivitiesForm();
                ClubPointList.ShowDialog();
            };
            //設定當選擇課程數量變動時
            Course.Instance.SelectedListChanged += delegate
            {
                Course.Instance.RibbonBarItems["資料統計"]["報表"]["社團相關報表"]["社團活動學生點名單"].Enable = (Course.Instance.SelectedList.Count > 0 && User.Acl["JHBehavior.Course.Ribbon0210"].Executable);
            }; 
            #endregion

            #region 設定權限
            Catalog ribbon = RoleAclSource.Instance["課程"]["報表"];
            ribbon.Add(new RibbonFeature("JHBehavior.Course.Ribbon0210", "社團活動學生點名單")); 
            #endregion


        }
    }
}