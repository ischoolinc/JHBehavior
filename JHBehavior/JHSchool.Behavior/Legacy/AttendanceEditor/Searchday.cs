using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using System.Xml;
using FISCA.DSAUtil;
using K12.Data;

namespace JHSchool.Behavior.Legacy.AttendanceEditor
{
    public partial class Searchday : BaseForm
    {
        K12.Data.Configuration.ConfigData cd;

        string _ConfigText;

        public Searchday(string ConfigText)
        {
            InitializeComponent();

            _ConfigText = ConfigText;

            cd = K12.Data.School.Configuration[_ConfigText];

            DSXmlHelper dsx = new DSXmlHelper("WeekSetup");
            foreach (ListViewItem each in listView1.Items)
            {
                dsx.AddElement("Day");
                dsx.SetAttribute("Day", "Detail", each.Text);
            }

            string cdIN = cd["星期設定"];

            XmlElement day;

            if (cdIN != "")
            {
                day = XmlHelper.LoadXml(cdIN);
            }
            else
            {
                day = dsx.BaseElement;
            }

            //if (day == null)
            //{
            //    DSXmlHelper dsx = new DSXmlHelper("WeekSetup");
            //    foreach(ListViewItem each in listView1.Items)
            //    {
            //        dsx.AddElement("Day");
            //        dsx.SetAttribute("Day", "Detail", each.Text);
            //    }

            //    day = dsx.BaseElement;
            //}

            List<string> list = new List<string>();
            foreach (XmlNode each in day.SelectNodes("Day"))
            {
                XmlElement each2 = each as XmlElement;
                list.Add(each2.GetAttribute("Detail"));
            }

            foreach (ListViewItem each in listView1.Items)
            {
                if (list.Contains(each.Text))
                {
                    each.Checked = true;
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            DSXmlHelper dsx = new DSXmlHelper("WeekSetup");

            foreach (ListViewItem each in listView1.Items)
            {
                if (each.Checked)
                {
                    dsx.AddElement("Day");
                    dsx.SetAttribute("Day", "Detail", each.Text);

                }
            }
            cd["星期設定"] = dsx.BaseElement.OuterXml;
            try
            {
                cd.Save();
            }
            catch(Exception ex)
            {
                MsgBox.Show("設定檔儲存錯誤" +ex.Message);
                return;
            }
            K12.Data.School.Configuration.Sync(_ConfigText);
            MsgBox.Show("星期設定儲存完成!");
            this.Close();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            foreach (ListViewItem each in listView1.Items)
            {
                each.Checked = checkBox1.Checked;
            }
        }
    }
}
