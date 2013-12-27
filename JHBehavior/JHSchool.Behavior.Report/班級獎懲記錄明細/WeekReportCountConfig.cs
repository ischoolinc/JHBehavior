using System;
using System.Windows.Forms;
using System.Xml;
using FISCA.DSAUtil;
using Framework;
using FISCA.Presentation.Controls;

namespace JHSchool.Behavior.Report
{
    public partial class WeekReportCountConfig : BaseForm
    {
        private string _reportName = "";

        /// <summary>
        /// 傳入設定字串
        /// </summary>
        /// <param name="reportname"></param>
        /// <param name="sizeIndex"></param>
        /// <param name="xml"></param>
        public WeekReportCountConfig(string reportname)
        {
            InitializeComponent();

            //設定檔
            _reportName = reportname;
            //設定畫面
            BingData();
        }

        private void BingData()
        {
            ConfigData cd = User.Configuration[_reportName];
            XmlElement config = cd.GetXml("XmlData", null);

            DSXmlHelper DSXML = new DSXmlHelper(config);
            comboBoxEx1.SelectedIndex = int.Parse(DSXML.GetAttribute("Print/@PaperSize"));
            checkBoxX1.Checked = bool.Parse(DSXML.GetAttribute("MeritDemerit/@Merit"));
            checkBoxX3.Checked = bool.Parse(DSXML.GetAttribute("MeritDemerit/@DemeritClear"));
            checkBoxX4.Checked = bool.Parse(DSXML.GetAttribute("MeritDemerit/@DemeritNotClear"));
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            #region 儲存 Preference

            //XmlElement config = CurrentUser.Instance.Preference[_reportName];
            ConfigData cd = User.Configuration[_reportName];
            XmlElement config = cd.GetXml("XmlData", null);

            DSXmlHelper DSXML = new DSXmlHelper(config);

            DSXML.SetAttribute("Print", "PaperSize", comboBoxEx1.SelectedIndex.ToString());

            DSXML.SetAttribute("MeritDemerit", "Merit", checkBoxX1.Checked.ToString());
            DSXML.SetAttribute("MeritDemerit", "DemeritClear", checkBoxX3.Checked.ToString());
            DSXML.SetAttribute("MeritDemerit", "DemeritNotClear", checkBoxX4.Checked.ToString());

            cd.SetXml("XmlData", DSXML.BaseElement);
            cd.Save();

            #endregion

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void checkBoxX1_CheckedChanged(object sender, EventArgs e)
        {
            CheckCheckBoxMode();
        }

        private void checkBoxX3_CheckedChanged(object sender, EventArgs e)
        {
            CheckCheckBoxMode();
        }

        private void checkBoxX4_CheckedChanged(object sender, EventArgs e)
        {
            CheckCheckBoxMode();
        }


        private void CheckCheckBoxMode()
        {
            if (!checkBoxX1.Checked && !checkBoxX3.Checked && !checkBoxX4.Checked)
            {
                FISCA.Presentation.Controls.MsgBox.Show("獎勵/懲戒/銷過...等3選項,至少必須選擇一項\n否則將會無法列印出任何資料!", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}