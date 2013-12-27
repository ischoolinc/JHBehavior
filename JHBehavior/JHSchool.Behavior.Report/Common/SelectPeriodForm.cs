using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using System.Xml;
using Framework;
using FISCA.Presentation.Controls;

namespace JHSchool.Behavior.Report
{
    public partial class SelectPeriodForm : BaseForm
    {
        private BackgroundWorker _BGWPeriodList;
        private string _key;

        private List<string> periodList = new List<string>();

        public SelectPeriodForm()
        {
            InitializeComponent();
        }

        public SelectPeriodForm(string key)
        {
            InitializeComponent();
            _key = key;
        }

        void _BGWPeriodList_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pictureBox1.Visible = false;

            foreach (string period in periodList)
            {
                ListViewItem item = new ListViewItem(period);
                listViewEx1.Items.Add(item);
            }

            #region 讀取列印設定 Preference

            //XmlElement config = CurrentUser.Instance.Preference["缺曠週報表_依節次統計_列印設定"];
            //XmlElement config = CurrentUser.Instance.Preference[_key];
            XmlElement config = User.Configuration[_key].GetXml("XmlData", null);

            if (config == null) return;

            foreach (XmlElement period in config.SelectNodes("Period"))
            {
                string name = period.GetAttribute("Name");
                foreach (ListViewItem item in listViewEx1.Items)
                {
                    if (item.Text == name)
                        item.Checked = true;
                }
            }

            #endregion
        }

        void _BGWPeriodList_DoWork(object sender, DoWorkEventArgs e)
        {
            foreach (XmlElement var in JHSchool.Compatibility.Feature.Basic.Config.GetPeriodList().GetContent().GetElements("Period"))
            {
                if (!periodList.Contains(var.GetAttribute("Name")))
                    periodList.Add(var.GetAttribute("Name"));
            }
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            #region 更新列印設定 Preference

            XmlElement config = User.Configuration[_key].GetXml("XmlData", null);

            if (config == null)
            {
                config = K12.Data.XmlHelper.LoadXml("<PeriodSetup/>");
            }

            config.RemoveAll();

            foreach (ListViewItem item in listViewEx1.Items)
            {
                XmlElement period = config.OwnerDocument.CreateElement("Period");
                period.SetAttribute("Name", item.Text);
                if (item.Checked == true)
                    config.AppendChild(period);
            }

            ConfigData cd = User.Configuration[_key];
            cd.SetXml("XmlData", config);
            cd.Save();

            #endregion

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void SelectPeriodForm_Load(object sender, EventArgs e)
        {
            if (!DesignMode)
            {

                _BGWPeriodList = new BackgroundWorker();
                _BGWPeriodList.DoWork += new DoWorkEventHandler(_BGWPeriodList_DoWork);
                _BGWPeriodList.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWPeriodList_RunWorkerCompleted);
                _BGWPeriodList.RunWorkerAsync();
                listViewEx1.Items.Clear();
            }
        }

        private void checkBoxX1_CheckedChanged(object sender, EventArgs e)
        {
            foreach (ListViewItem each in listViewEx1.Items)
            {
                each.Checked = checkBoxX1.Checked;
            }
        }
    }
}