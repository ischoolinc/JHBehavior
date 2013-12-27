using System;
using System.Windows.Forms;
using System.Xml;
using Framework;
using FISCA.Presentation.Controls;
namespace JHSchool.Behavior.Report
{
    public partial class WeekAbsenceReportConfig : BaseForm
    {
        private string _reportName = "";
        public WeekAbsenceReportConfig(string reportname, int sizeIndex, bool CheckClass, bool CheckWeek)
        {
            InitializeComponent();

            _reportName = reportname;
            comboBoxEx1.SelectedIndex = sizeIndex;
            checkBoxX1.Checked = CheckClass;
            checkBoxX2.Checked = CheckWeek;
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            #region Àx¦s Preference

            //XmlElement config = CurrentUser.Instance.Preference[_reportName];
            ConfigData cd = User.Configuration[_reportName];
            XmlElement config = cd.GetXml("XmlData", null);

            if (config == null)
                config = new XmlDocument().CreateElement(_reportName);

            #region PaperSize
            XmlElement print = config.OwnerDocument.CreateElement("Print");
            print.SetAttribute("PaperSize", comboBoxEx1.SelectedIndex.ToString());

            if (config.SelectSingleNode("Print") == null)
                config.AppendChild(print);
            else
                config.ReplaceChild(print, config.SelectSingleNode("Print")); 
            #endregion

            #region CheckClass
            XmlElement TestClass = config.OwnerDocument.CreateElement("CheckClass");
            TestClass.SetAttribute("Class", checkBoxX1.Checked.ToString());

            if (config.SelectSingleNode("CheckClass") == null)
                config.AppendChild(TestClass);
            else
                config.ReplaceChild(TestClass, config.SelectSingleNode("CheckClass")); 
            #endregion

            #region CheckWeek
            XmlElement TestWeek = config.OwnerDocument.CreateElement("CheckWeek");
            TestWeek.SetAttribute("Week", checkBoxX2.Checked.ToString());

            if (config.SelectSingleNode("CheckWeek") == null)
                config.AppendChild(TestWeek);
            else
                config.ReplaceChild(TestWeek, config.SelectSingleNode("CheckWeek"));
            #endregion

            //CurrentUser.Instance.Preference[_reportName] = config;

            cd.SetXml("XmlData", config);
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
    }
}