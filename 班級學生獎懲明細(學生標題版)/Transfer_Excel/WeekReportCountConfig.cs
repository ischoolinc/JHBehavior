using FISCA.DSAUtil;
using FISCA.Presentation.Controls;
using K12.Data.Configuration;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace Transfer_Excel
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
            ConfigData cd = K12.Data.School.Configuration[_reportName];
            XmlElement config = cd.GetXml("XmlData", null);

            DSXmlHelper DSXML = new DSXmlHelper(config);
            comboBoxEx1.SelectedIndex = int.Parse(DSXML.GetAttribute("Print/@PaperSize"));
            checkBoxX1.Checked = bool.Parse(DSXML.GetAttribute("MeritDemerit/@Merit"));
            checkBoxX2.Checked = bool.Parse(DSXML.GetAttribute("MeritDemerit/@DemeritClear"));
            checkBoxX3.Checked = bool.Parse(DSXML.GetAttribute("MeritDemerit/@DemeritNotClear"));
        }

        private void WeekReportCountConfig_Load(object sender, EventArgs e)
        {

        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            #region 儲存 Preference

            //XmlElement config = CurrentUser.Instance.Preference[_reportName];
            ConfigData cd = K12.Data.School.Configuration[_reportName];
            XmlElement config = cd.GetXml("XmlData", null);

            DSXmlHelper DSXML = new DSXmlHelper(config);

            DSXML.SetAttribute("Print", "PaperSize", comboBoxEx1.SelectedIndex.ToString());
            DSXML.SetAttribute("MeritDemerit", "Merit", checkBoxX1.Checked.ToString());
            DSXML.SetAttribute("MeritDemerit", "DemeritClear", checkBoxX2.Checked.ToString());
            DSXML.SetAttribute("MeritDemerit", "DemeritNotClear", checkBoxX3.Checked.ToString());

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

        private void checkBoxX2_CheckedChanged(object sender, EventArgs e)
        {
            CheckCheckBoxMode();
        }

        private void checkBoxX3_CheckedChanged(object sender, EventArgs e)
        {
            CheckCheckBoxMode();
        }

        private void CheckCheckBoxMode()
        {
            if (!checkBoxX1.Checked && !checkBoxX2.Checked && !checkBoxX3.Checked)
            {
                FISCA.Presentation.Controls.MsgBox.Show("獎勵/懲戒/銷過...等3選項,至少必須選擇一項\n否則將會無法列印出任何資料!", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
