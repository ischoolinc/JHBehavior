using System;
using System.Windows.Forms;
using System.Xml;
using Framework;

namespace JHSchool.Behavior.Report
{
    public partial class AbsenceDetailForm : SelectDateRangeForm
    {
        private int _sizeIndex = 0;
        private string _reportstyle;
        public int PaperSize
        {
            get { return _sizeIndex; }
        }
        
        public string ReportStyle
        {
            get { return _reportstyle; }
        }

        public AbsenceDetailForm()
        {
            InitializeComponent();
            LoadPreference();
        }

        private void LoadPreference()
        {
            #region Åª¨ú Preference

            //XmlElement config = CurrentUser.Instance.Preference["¯Z¯Å¯ÊÃm©ú²Óªí"];
            ConfigData cd = User.Configuration["¯Z¯Å¯ÊÃm°O¿ý©ú²Ó"];
            XmlElement config = cd.GetXml("XmlData", null);

            if (config != null)
            {
                XmlElement print = (XmlElement)config.SelectSingleNode("Print");

                if (print != null)
                {
                    if (print.HasAttribute("PaperSize"))
                        _sizeIndex = int.Parse(print.GetAttribute("PaperSize"));
                    if (print.HasAttribute("ReportStyle"))
                        _reportstyle = print.GetAttribute("ReportStyle");
                }
                else
                {
                    XmlElement newPrint = config.OwnerDocument.CreateElement("Print");
                    newPrint.SetAttribute("PaperSize", "0");
                    config.AppendChild(newPrint);
                    //CurrentUser.Instance.Preference["¯Z¯Å¯ÊÃm°O¿ý©ú²Ó"] = config;
                    cd.SetXml("XmlData", config);
                }
            }
            else
            {
                #region ²£¥ÍªÅ¥Õ³]©wÀÉ
                config = new XmlDocument().CreateElement("¯Z¯Å¯ÊÃm°O¿ý©ú²Ó");
                XmlElement printSetup = config.OwnerDocument.CreateElement("Print");
                printSetup.SetAttribute("PaperSize", "0");
                config.AppendChild(printSetup);
                //CurrentUser.Instance.Preference["¯Z¯Å¯ÊÃm°O¿ý©ú²Ó"] = config;
                cd.SetXml("XmlData", config);
                #endregion
            }

            cd.Save();

            #endregion
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //-----
            AttendanceConfig config = new AttendanceConfig("¯Z¯Å¯ÊÃm°O¿ý©ú²Ó", _sizeIndex, _reportstyle);
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