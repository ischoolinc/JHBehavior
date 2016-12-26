using Aspose.Cells;
using FISCA.DSAUtil;
using FISCA.Presentation.Controls;
using K12.Data;
using K12.Data.Configuration;
using K12.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace Transfer_Excel
{
    public partial class DisciplineDetailForm : SelectDateRangeForm
    {
        private int _sizeIndex = 0;
        public int PaperSize
        {
            get { return _sizeIndex; }
        }

        public Dictionary<string, bool> SetupDic = new Dictionary<string, bool>();

        public DisciplineDetailForm()
        {
            InitializeComponent();

            LoadPreference();
        }

        private void LoadPreference()
        {
            #region 讀取 Preference
            SetupDic.Clear();
            ConfigData cd = K12.Data.School.Configuration["班級獎懲記錄明細"];
            XmlElement config = cd.GetXml("XmlData", null);
            //XmlElement config = CurrentUser.Instance.Preference["班級獎懲明細表"];

            if (config != null)
            {
                DSXmlHelper DSXML = new DSXmlHelper(config);

                _sizeIndex = int.Parse(DSXML.GetAttribute("Print/@PaperSize"));

                if (DSXML.GetElement("MeritDemerit") == null) //處理各校目前設定值
                {
                    DSXML.AddElement("MeritDemerit");
                    DSXML.SetAttribute("MeritDemerit", "Merit", "True");
                    DSXML.SetAttribute("MeritDemerit", "DemeritClear", "True");
                    DSXML.SetAttribute("MeritDemerit", "DemeritNotClear", "True");
                    cd.SetXml("XmlData", DSXML.BaseElement);
                    cd.Save();
                }

                SetupDic.Add("獎勵", bool.Parse(DSXML.GetAttribute("MeritDemerit/@Merit")));
                SetupDic.Add("懲戒未銷過", bool.Parse(DSXML.GetAttribute("MeritDemerit/@DemeritClear")));
                SetupDic.Add("懲戒銷過", bool.Parse(DSXML.GetAttribute("MeritDemerit/@DemeritNotClear")));

            }
            else
            {
                //使用預設
                DSXmlHelper DSXML = new DSXmlHelper("XmlData");
                DSXML.AddElement("Print");
                DSXML.SetAttribute("Print", "PaperSize", "0");
                DSXML.AddElement("MeritDemerit");
                DSXML.SetAttribute("MeritDemerit", "Merit", "True");
                DSXML.SetAttribute("MeritDemerit", "DemeritClear", "True");
                DSXML.SetAttribute("MeritDemerit", "DemeritNotClear", "True");
                cd.SetXml("XmlData", DSXML.BaseElement);
                cd.Save();

                SetupDic.Add("獎勵", bool.Parse(DSXML.GetAttribute("MeritDemerit/@Merit")));
                SetupDic.Add("懲戒未銷過", bool.Parse(DSXML.GetAttribute("MeritDemerit/@DemeritClear")));
                SetupDic.Add("懲戒銷過", bool.Parse(DSXML.GetAttribute("MeritDemerit/@DemeritNotClear")));
            }
            #endregion
        }

        private void DisciplineDetailForm_Load(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            WeekReportCountConfig config = new WeekReportCountConfig("班級獎懲記錄明細");
            if (config.ShowDialog() == DialogResult.OK)
            {
                LoadPreference();
            }
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }


    }
}
