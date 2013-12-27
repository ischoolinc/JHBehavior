using System;
using System.Windows.Forms;
using System.Xml;
using Framework;

namespace JHSchool.Behavior.Report
{
    public partial class AbsenceDetailForm : SelectDateRangeForm
    {
        private int _sizeIndex = 0;

        public int PaperSize
        {
            get { return _sizeIndex; }
        }

        public AbsenceDetailForm()
        {
            InitializeComponent();
            LoadPreference();
        }

        private void LoadPreference()
        {
            #region 讀取 Preference

            //XmlElement config = CurrentUser.Instance.Preference["班級缺曠明細表"];
            ConfigData cd = User.Configuration["班級缺曠記錄明細"];
            XmlElement config = cd.GetXml("XmlData", null);

            if (config != null)
            {
                XmlElement print = (XmlElement)config.SelectSingleNode("Print");

                if (print != null)
                {
                    if (print.HasAttribute("PaperSize"))
                        _sizeIndex = int.Parse(print.GetAttribute("PaperSize"));
                }
                else
                {
                    XmlElement newPrint = config.OwnerDocument.CreateElement("Print");
                    newPrint.SetAttribute("PaperSize", "0");
                    config.AppendChild(newPrint);
                    //CurrentUser.Instance.Preference["班級缺曠記錄明細"] = config;
                    cd.SetXml("XmlData", config);
                }
            }
            else
            {
                #region 產生空白設定檔
                config = new XmlDocument().CreateElement("班級缺曠記錄明細");
                XmlElement printSetup = config.OwnerDocument.CreateElement("Print");
                printSetup.SetAttribute("PaperSize", "0");
                config.AppendChild(printSetup);
                //CurrentUser.Instance.Preference["班級缺曠記錄明細"] = config;
                cd.SetXml("XmlData", config);
                #endregion
            }

            cd.Save();

            #endregion
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            AttendanceConfig config = new AttendanceConfig("班級缺曠記錄明細", _sizeIndex);
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