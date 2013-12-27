using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Permission;
using FISCA.Presentation;
using K12.Data;

namespace K12.Behavior.Address
{
    public static class Program
    {
        [MainMethod()]
        static public void Main()
        {
            RibbonBarItem classSpecialItem = K12.Presentation.NLDPanels.Class.RibbonBarItems["學務"];
            classSpecialItem["聯絡資訊管理"].Size = RibbonBarButton.MenuButtonSize.Medium;
            classSpecialItem["聯絡資訊管理"].Image = Properties.Resources.home_write_64;
            classSpecialItem["聯絡資訊管理"].Enable = false;
            classSpecialItem["聯絡資訊管理"].Click += delegate
            {
                AddressEditForm address = new AddressEditForm();
                address.ShowDialog();
            };

            K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
            {
                classSpecialItem["聯絡資訊管理"].Enable = (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0 && Permissions.聯絡資訊管理權限);
            };

            Catalog ribbon = RoleAclSource.Instance["班級"]["功能按鈕"];
            ribbon.Add(new RibbonFeature("JHBehavior.Class.Ribbon0210", "聯絡資訊管理"));
        }
    }
}
