using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using DevComponents.DotNetBar.Controls;
using DevComponents.DotNetBar.Rendering;
using Framework;
using FISCA.Presentation.Controls;

namespace JHSchool.Behavior.Report
{
    public partial class DemeritConfigForm : BaseForm
    {
        private byte[] _buffer = null;
        private string base64 = null;
        private bool _isUpload = false;
        private bool _defaultTemplate;
        private bool _printHasRecordOnly;
        private DateRangeModeNew _mode = DateRangeModeNew.Month;
        private bool _printStudentList;
        private bool _printRemark; //�O�_�C�L�Ƶ�

        public DemeritConfigForm(bool defaultTemplate, DateRangeModeNew mode, byte[] buffer,
            string name, string address, string condName, string condNumber, 
            bool printStudentList, bool printRemark)
        {
            InitializeComponent();
            #region �p�G�t�Ϊ�Renderer�OOffice2007Renderer�A�P��_ClassTeacherView,_CategoryView���C��
            if (GlobalManager.Renderer is Office2007Renderer)
            {
                ((Office2007Renderer)GlobalManager.Renderer).ColorTableChanged += new EventHandler(ScoreCalcRuleEditor_ColorTableChanged);
                SetForeColor(this);
            }
            #endregion
            _defaultTemplate = defaultTemplate;
            _mode = mode;
            _printStudentList = printStudentList;
            _printRemark = printRemark;

            if (buffer != null)
                _buffer = buffer;

            if (defaultTemplate)
                radioButton1.Checked = true;
            else
                radioButton2.Checked = true;

            cbPrintStudentList.Checked = printStudentList;
            cbPrintRemark.Checked = printRemark;
            switch (mode)
            {
                case DateRangeModeNew.Month:
                    radioButton3.Checked = true;
                    break;
                case DateRangeModeNew.Week:
                    radioButton4.Checked = true;
                    break;
                case DateRangeModeNew.Custom:
                    radioButton5.Checked = true;
                    break;
                default:
                    throw new Exception("Date Range Mode Error.");
            }

            //�]�w ComboBox
            Dictionary<ComboBoxEx, string> cboBoxes = new Dictionary<ComboBoxEx, string>();
            cboBoxes.Add(comboBoxEx1, name);
            cboBoxes.Add(comboBoxEx2, address);
            cboBoxes.Add(comboBoxEx4, condName);

            foreach (ComboBoxEx var in cboBoxes.Keys)
            {
                var.SelectedIndex = 0;
                foreach (DevComponents.Editors.ComboItem item in var.Items)
                {
                    if (item.Text == cboBoxes[var])
                    {
                        var.SelectedIndex = var.Items.IndexOf(item);
                        break;
                    }
                }
            }

            //�]�w NumericUpDown
            decimal tryValue;
            if (condNumber == "0")
                condNumber = "1";
            numericUpDown1.Value = (decimal.TryParse(condNumber, out tryValue)) ? tryValue : 1;           
        }

        void ScoreCalcRuleEditor_ColorTableChanged(object sender, EventArgs e)
        {
            SetForeColor(this);
        }

        private void SetForeColor(Control parent)
        {
            foreach (Control var in parent.Controls)
            {
                if (var is RadioButton)
                    var.ForeColor = ((Office2007Renderer)GlobalManager.Renderer).ColorTable.CheckBoxItem.Default.Text;
                SetForeColor(var);
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                radioButton2.Checked = false;
                _defaultTemplate = true;
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                radioButton1.Checked = false;
                _defaultTemplate = false;
            }
        }

        private void checkBoxX2_CheckedChanged(object sender, EventArgs e)
        {
            _printStudentList = cbPrintStudentList.Checked;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "�t�s�s��";
            sfd.FileName = "�g�ٳq����.doc";
            sfd.Filter = "Word�ɮ� (*.doc)|*.doc|�Ҧ��ɮ� (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(ProjectResource.�g�ٳq����, 0, ProjectResource.�g�ٳq����.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("���w���|�L�k�s���C", "�t�s�ɮץ���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "�t�s�s��";
            sfd.FileName = "�ۭq�g�ٳq����.doc";
            sfd.Filter = "Word�ɮ� (*.doc)|*.doc|�Ҧ��ɮ� (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    if (Aspose.Words.Document.DetectFileFormat(new MemoryStream(_buffer)) == Aspose.Words.LoadFormat.Doc)
                        fs.Write(_buffer, 0, _buffer.Length);
                    else
                        fs.Write(ProjectResource.�g�ٳq����, 0, ProjectResource.�g�ٳq����.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("���w���|�L�k�s���C", "�t�s�ɮץ���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "��ܦۭq���g�ٳq����d��";
            ofd.Filter = "Word�ɮ� (*.doc)|*.doc";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if (Aspose.Words.Document.DetectFileFormat(ofd.FileName) == Aspose.Words.LoadFormat.Doc)
                    {
                        FileStream fs = new FileStream(ofd.FileName, FileMode.Open);

                        byte[] tempBuffer = new byte[fs.Length];
                        fs.Read(tempBuffer, 0, tempBuffer.Length);
                        base64 = Convert.ToBase64String(tempBuffer);
                        _isUpload = true;
                        fs.Close();
                        FISCA.Presentation.Controls.MsgBox.Show("�W�Ǧ��\�C");
                    }
                    else
                        FISCA.Presentation.Controls.MsgBox.Show("�W���ɮ׮榡����");
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("���w���|�L�k�s���C", "�}���ɮץ���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            #region �x�s Preference

            ConfigData cd = User.Configuration["�g�ٳq����"];
            XmlElement config = cd.GetXml("XmlData", null);

            //XmlElement config = CurrentUser.Instance.Preference["�g�ٳq����"];

            if (config == null)
            {
                config = new XmlDocument().CreateElement("�g�ٳq����");
            }

            config.SetAttribute("Default", _defaultTemplate.ToString());

            XmlElement customize = config.OwnerDocument.CreateElement("CustomizeTemplate");
            XmlElement mode = config.OwnerDocument.CreateElement("DateRangeMode");
            XmlElement receive = config.OwnerDocument.CreateElement("Receive");
            XmlElement conditions = config.OwnerDocument.CreateElement("Conditions");
            XmlElement PrintStudentList = config.OwnerDocument.CreateElement("PrintStudentList");
            XmlElement PrintRemark = config.OwnerDocument.CreateElement("PrintRemark");

            PrintStudentList.SetAttribute("Checked", _printStudentList.ToString());
            config.ReplaceChild(PrintStudentList, config.SelectSingleNode("PrintStudentList"));

            PrintRemark.SetAttribute("Checked", _printRemark.ToString());
            config.ReplaceChild(PrintRemark, config.SelectSingleNode("PrintRemark"));

            if (_isUpload)
            {
                customize.InnerText = base64;
                config.ReplaceChild(customize, config.SelectSingleNode("CustomizeTemplate"));
            }

            mode.InnerText = ((int)_mode).ToString();
            config.ReplaceChild(mode, config.SelectSingleNode("DateRangeMode"));

            receive.SetAttribute("Name", ((DevComponents.Editors.ComboItem)comboBoxEx1.SelectedItem).Text);
            receive.SetAttribute("Address", ((DevComponents.Editors.ComboItem)comboBoxEx2.SelectedItem).Text);
            if (config.SelectSingleNode("Receive") == null)
                config.AppendChild(receive);
            else
                config.ReplaceChild(receive, config.SelectSingleNode("Receive"));

            conditions.SetAttribute("ConditionName", ((DevComponents.Editors.ComboItem)comboBoxEx4.SelectedItem).Text);
            conditions.SetAttribute("ConditionNumber", numericUpDown1.Value.ToString());
            if (config.SelectSingleNode("Conditions") == null)
                config.AppendChild(conditions);
            else
                config.ReplaceChild(conditions, config.SelectSingleNode("Conditions"));

            cd.SetXml("XmlData", config);
            cd.Save();

            //CurrentUser.Instance.Preference["�g�ٳq����"] = config;

            #endregion

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked)
            {
                radioButton4.Checked = false;
                radioButton5.Checked = false;
                _mode = DateRangeModeNew.Month;
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked)
            {
                radioButton3.Checked = false;
                radioButton5.Checked = false;
                _mode = DateRangeModeNew.Week;
            }
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked)
            {
                radioButton3.Checked = false;
                radioButton4.Checked = false;
                _mode = DateRangeModeNew.Custom;
            }
        }

        private void comboBoxEx4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxEx4.SelectedIndex == 0)
            {
                numericUpDown1.Enabled = false;
                numericUpDown1.Value = 1;
            }
            else
            {
                numericUpDown1.Enabled = true;
            }
        }

        private void cbPrintRemark_CheckedChanged(object sender, EventArgs e)
        {
            _printRemark = cbPrintRemark.Checked;
        }
    }
}