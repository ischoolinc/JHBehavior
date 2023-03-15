using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using DevComponents.DotNetBar;
using Framework;
using SmartSchool.Common;
using Framework.Feature;
using JHSchool.Editor;
using JHSchool.Behavior.Editor;
using JHSchool.Behavior.Legacy;
using Framework.Legacy;
using JHSchool.Behavior.Feature;
using Framework.Security;
using Framework.DotNetBar;
using JHSchool.Data;
using FISCA.LogAgent;

namespace JHSchool.Behavior.StudentExtendControls
{
    partial class AttendanceForm : FISCA.Presentation.Controls.BaseForm
    {
        private ErrorProvider _errorProvider = new ErrorProvider();
        private EditorStatus _status;

        private JHAttendanceRecord _editor;
        private List<K12.Data.AbsenceMappingInfo> _absenceList;
        private List<K12.Data.PeriodMappingInfo> _periodList;
        private List<K12.Data.AbsenceMappingInfo> absenceList;
        private K12.Data.AbsenceMappingInfo _checkedAbsence;

        //Log�ϥ�
        //���e
        private Dictionary<string, string> DicBeforeLog = new Dictionary<string, string>();
        //����
        private Dictionary<string, string> DicAfterLog = new Dictionary<string, string>();

        /// <summary>
        /// ���m�Ǵ��έp - �ϥ�
        /// </summary>
        public AttendanceForm(EditorStatus status, JHAttendanceRecord editor, List<K12.Data.PeriodMappingInfo> periodList, FeatureAce permission, int SchoolYear, int Semester)
        {
            InitializeComponent();

            #region ��l�ƾǦ~�׾Ǵ�

            //�Ǧ~��
            intSchoolYear.Value = SchoolYear;
            intSemester.Value = Semester;

            #endregion

            #region ��l�Ư��m���O

            //��l�Ʈ�,�Y���o�̷s���m���
            absenceList = K12.Data.AbsenceMapping.SelectAll();

            foreach (K12.Data.AbsenceMappingInfo info in absenceList)
            {
                RadioButton rb = new RadioButton();
                //���m�O,�Y�g,����
                rb.Text = info.Name + "(" + info.HotKey.ToUpper() + ")";
                rb.AutoSize = true;
                rb.Font = new Font(SmartSchool.Common.FontStyles.GeneralFontFamily, 9.25f);
                rb.Tag = info;
                rb.CheckedChanged += delegate(object sender, EventArgs e)
                {
                    if (rb.Checked)
                    {
                        foreach (DataGridViewCell cell in dataGridViewX1.SelectedCells)
                        {
                            cell.Value = (rb.Tag as K12.Data.AbsenceMappingInfo).Abbreviation;
                        }
                    }
                };
                flpAbsence.Controls.Add(rb);
            }

            //��Ĥ@�ӯ��m���O�]���w�]��
            RadioButton fouse = flpAbsence.Controls[0] as RadioButton;
            fouse.Checked = true;

            #endregion

            #region ��l�Ƹ`����

            foreach (K12.Data.PeriodMappingInfo info in periodList)
            {
                //Log�ϥ�
                if (!DicBeforeLog.ContainsKey(info.Name))
                {
                    DicBeforeLog.Add(info.Name, "");
                }
                if (!DicAfterLog.ContainsKey(info.Name))
                {
                    DicAfterLog.Add(info.Name, "");
                }

                DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
                column.Width = 40;
                column.HeaderText = info.Name;
                column.Tag = info;
                dataGridViewX1.Columns.Add(column);

                //PeriodControl pc = new PeriodControl();
                //pc.Label.Text = info.Name;
                //pc.Tag = info;
                //pc.TextBox.KeyUp += delegate(object sender, KeyEventArgs e)
                //{
                //    var txtBox = sender as DevComponents.DotNetBar.Controls.TextBoxX;
                //    foreach (AbsenceMappingInfo absenceInfo in absenceList)
                //    {
                //        if (KeyConverter.GetKeyMapping(e) == absenceInfo.HotKey || KeyConverter.GetKeyMapping(e) == absenceInfo.HotKey.ToUpper())
                //        {
                //            txtBox.Text = absenceInfo.Abbreviation;
                //            if (flpPeriod.GetNextControl(pc, true) != null)
                //                (flpPeriod.GetNextControl(pc, true) as PeriodControl).TextBox.Focus();
                //            return;
                //        }
                //    }
                //    txtBox.SelectAll();
                //};
                //pc.TextBox.MouseDoubleClick += new MouseEventHandler(TextBox_MouseDoubleClick);
                //flpPeriod.Controls.Add(pc);
            }


            DataGridViewRow row = new DataGridViewRow();
            row.CreateCells(dataGridViewX1);
            dataGridViewX1.Rows.Add(row);
            dataGridViewX1.AutoResizeColumns();
            #endregion

            dateTimeInput1.Value = DateTime.Today;

            btnSave.Visible = permission.Editable;
            intSchoolYear.Enabled = permission.Editable;
            intSemester.Enabled = permission.Editable;
            panelAbsence.Enabled = permission.Editable;
            //pancelAttendence.Enabled = permission.Editable;
            dataGridViewX1.Enabled = permission.Editable;

            if (status == EditorStatus.Insert)
            {
                Text = "�޲z�ǥͯ��m�����i�s�W�Ҧ��j";
            }

            _status = status;
            _editor = editor;
            _absenceList = absenceList;
            _periodList = periodList;
        }

        /// <summary>
        /// ��s���m���
        /// </summary>
        public AttendanceForm(EditorStatus status, JHAttendanceRecord editor, List<K12.Data.PeriodMappingInfo> periodList, FeatureAce permission)
        {
            InitializeComponent();

            #region ��l�ƾǦ~�׾Ǵ�

            //�Ǧ~��
            intSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
            intSemester.Value = int.Parse(K12.Data.School.DefaultSemester);

            #endregion

            #region ��l�Ư��m���O

            //��l�Ʈ�,�Y���o�̷s���m���
            absenceList = K12.Data.AbsenceMapping.SelectAll();

            foreach (K12.Data.AbsenceMappingInfo info in absenceList)
            {
                RadioButton rb = new RadioButton();
                rb.Text = info.Name + "(" + info.HotKey.ToUpper() + ")";
                rb.AutoSize = true;
                rb.Font = new Font(SmartSchool.Common.FontStyles.GeneralFontFamily, 9.25f);
                rb.Tag = info;

                rb.Click += delegate
                {
                    if (rb.Checked)
                    {
                        foreach (DataGridViewCell cell in dataGridViewX1.SelectedCells)
                        {
                            cell.Value = ((K12.Data.AbsenceMappingInfo)rb.Tag).Abbreviation;
                        }
                    }
                };
                //rb.CheckedChanged += delegate(object sender, EventArgs e)
                //{
                //    if (rb.Checked)
                //    {
                //        foreach (DataGridViewCell cell in dataGridViewX1.SelectedCells)
                //        {
                //            cell.Value = (rb.Tag as K12.Data.AbsenceMappingInfo).Abbreviation;
                //        }
                //    }
                //};
                flpAbsence.Controls.Add(rb);
            }

            //��Ĥ@�ӯ��m���O�]���w�]��
            RadioButton fouse = flpAbsence.Controls[0] as RadioButton;
            fouse.Checked = true;

            #endregion
            List<string> plist = new List<string>();
            #region ��l�Ƹ`����
            foreach (K12.Data.PeriodMappingInfo info in periodList)
            {
                //Log�ϥ�
                DicBeforeLog.Add(info.Name, "");
                DicAfterLog.Add(info.Name, "");

                DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
                column.Width = 40;
                column.HeaderText = info.Name;
                plist.Add(info.Name);
                column.Tag = info;
                dataGridViewX1.Columns.Add(column);

                //PeriodControl pc = new PeriodControl();
                //pc.Label.Text = info.Name;
                //pc.Tag = info;
                //pc.TextBox.KeyUp += delegate(object sender, KeyEventArgs e)
                //{
                //    var txtBox = sender as DevComponents.DotNetBar.Controls.TextBoxX;
                //    foreach (AbsenceMappingInfo absenceInfo in absenceList)
                //    {
                //        if (KeyConverter.GetKeyMapping(e) == absenceInfo.HotKey || KeyConverter.GetKeyMapping(e) == absenceInfo.HotKey.ToUpper())
                //        {
                //            txtBox.Text = absenceInfo.Abbreviation;
                //            if (flpPeriod.GetNextControl(pc, true) != null)
                //                (flpPeriod.GetNextControl(pc, true) as PeriodControl).TextBox.Focus();
                //            return;
                //        }
                //    }
                //    txtBox.SelectAll();
                //};
                //pc.TextBox.MouseDoubleClick += new MouseEventHandler(TextBox_MouseDoubleClick);
                //flpPeriod.Controls.Add(pc);
            }

            Campus.Windows.DataGridViewImeDecorator dec = new Campus.Windows.DataGridViewImeDecorator(this.dataGridViewX1, plist);

            DataGridViewRow row = new DataGridViewRow();
            row.CreateCells(dataGridViewX1);
            dataGridViewX1.Rows.Add(row);
            dataGridViewX1.AutoResizeColumns();
            #endregion

            dateTimeInput1.Value = DateTime.Today;

            btnSave.Visible = permission.Editable;
            intSchoolYear.Enabled = permission.Editable;
            intSemester.Enabled = permission.Editable;
            panelAbsence.Enabled = permission.Editable;
            //pancelAttendence.Enabled = permission.Editable;

            if (status == EditorStatus.Insert)
            {
                Text = "�޲z�ǥͯ��m�����i�s�W�Ҧ��j";
            }

            if (status == EditorStatus.Update)
            {
                Text = "�޲z�ǥͯ��m�����i�ק�Ҧ��j";

                dateTimeInput1.Value = editor.OccurDate;
                dateTimeInput1.Enabled = false;
                intSchoolYear.Text = editor.SchoolYear.ToString();
                intSemester.Text = editor.Semester.ToString();

                foreach (K12.Data.AttendancePeriod period in editor.PeriodDetail)
                {
                    #region Log

                    if (DicBeforeLog.ContainsKey(period.Period))
                    {
                        DicBeforeLog[period.Period] = period.AbsenceType;
                    }

                    #endregion

                    foreach (DataGridViewCell cell in dataGridViewX1.Rows[0].Cells)
                    {
                        K12.Data.PeriodMappingInfo info = cell.OwningColumn.Tag as K12.Data.PeriodMappingInfo;

                        if (info == null) continue;
                        if (period.Period != info.Name) continue;
                        if (period.AbsenceType == null) continue;

                        foreach (K12.Data.AbsenceMappingInfo ai in absenceList)
                        {
                            if (ai.Name != period.AbsenceType) continue;

                            cell.Value = ai.Abbreviation;
                            break;
                        }
                    }
                }
            }

            if (!permission.Editable)
                Text = "�޲z�ǥͯ��m�����i�˵��Ҧ��j";

            _status = status;
            _editor = editor;
            _absenceList = absenceList;
            _periodList = periodList;
        }

        //private string chengDateTime(DateTime x)
        //{
        //    string time = x.ToString();
        //    int y = time.IndexOf(' ');
        //    return time.Remove(y);
        //}

        private void btnSave_Click(object sender, EventArgs e)
        {

            //2023/3/14 - �W�[���ҨϥΪ̬O�_����J�ɶ�
            if (dateTimeInput1.Text == "0001/01/01 00:00:00" || dateTimeInput1.Text == "")
            {
                _errorProvider.SetError(dateTimeInput1, "�п�J�ɶ����");
                return;
            }
            else
            {
                _errorProvider.SetError(dateTimeInput1, "");
            }

            if (_status == EditorStatus.Insert)
            {
                foreach (JHAttendanceRecord each in JHAttendance.SelectByStudentIDs(new string[] { _editor.RefStudentID }))
                {
                    if (each.OccurDate == dateTimeInput1.Value)
                    {
                        _errorProvider.SetError(dateTimeInput1, "������w�������s�b�A�Ч�έק�Ҧ�");
                        return;
                    }
                }
            }

            _editor.OccurDate = dateTimeInput1.Value;
            _editor.SchoolYear = intSchoolYear.Value;
            _editor.Semester = intSemester.Value;

            List<K12.Data.AttendancePeriod> periodDetail = new List<K12.Data.AttendancePeriod>();

            foreach (DataGridViewCell cell in dataGridViewX1.Rows[0].Cells)
            {
                if (!string.IsNullOrEmpty(("" + cell.Value).Trim()))
                {
                    K12.Data.AttendancePeriod ap = new K12.Data.AttendancePeriod();

                    foreach (K12.Data.AbsenceMappingInfo ai in _absenceList)
                    {
                        if (ai.Abbreviation == ("" + cell.Value).Trim())
                        {
                            ap.Period = (cell.OwningColumn.Tag as K12.Data.PeriodMappingInfo).Name;
                            ap.AbsenceType = ai.Name;
                            //ap.AttendanceType = "�@��";

                        }
                    }

                    periodDetail.Add(ap);
                }
            }

            _editor.PeriodDetail = periodDetail;

            if (periodDetail.Count == 0)
            {
                FISCA.Presentation.Controls.MsgBox.Show("���i�x�s�L���m���e�����!");
                return;
            }

            if (_status == EditorStatus.Insert)
            {
                #region Insert
                try
                {
                    JHAttendance.Insert(_editor);
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("�s�W���m��ƥ���!" + ex.Message);
                    return;
                }

                StringBuilder sb = new StringBuilder();
                sb.Append("�ǥ͡u" + _editor.Student.Name + "�v");
                sb.Append("����u" + _editor.OccurDate.ToShortDateString() + "�v");
                sb.AppendLine("�s�W�@�����m��ơC");
                sb.AppendLine("�ԲӸ�ơG");
                foreach (K12.Data.AttendancePeriod each in _editor.PeriodDetail)
                {
                    sb.AppendLine("�`���u" + each.Period + "�v�]���u" + each.AbsenceType + "�v");
                }

                ApplicationLog.Log("�ǰȨt��.���m���", "�s�W�ǥͯ��m���", "student", _editor.Student.ID, sb.ToString());
                FISCA.Presentation.Controls.MsgBox.Show("�s�W�ǥͯ��m��Ʀ��\!");
                #endregion
            }
            else if (_status == EditorStatus.Update)
            {
                #region Update
                try
                {
                    JHAttendance.Update(_editor);
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("�ק���m��ƥ���!" + ex.Message);
                    return;
                }

                StringBuilder sb = new StringBuilder();
                sb.Append("�ǥ͡u" + _editor.Student.Name + "�v");
                sb.Append("����u" + _editor.OccurDate.ToShortDateString() + "�v");
                sb.AppendLine("���m��Ƥw�ק�C");
                sb.AppendLine("�ԲӸ�ơG");

                foreach (K12.Data.AttendancePeriod each in _editor.PeriodDetail)
                {
                    if (DicAfterLog.ContainsKey(each.Period))
                    {
                        DicAfterLog[each.Period] = each.AbsenceType;
                    }
                }

                bool LogMode = false;
                foreach (K12.Data.PeriodMappingInfo each in _periodList)
                {
                    if ( (DicAfterLog[each.Name] != DicBeforeLog[each.Name]))
                    {
                        sb.AppendLine("�`���u" + each.Name + "�v�ѡu" + DicBeforeLog[each.Name] + "�v�ܧ󬰡u" + DicAfterLog[each.Name] + "�v");
                        LogMode = true;
                    }
                }

                if (LogMode)
                {
                    ApplicationLog.Log("�ǰȨt��.���m���", "�ק�ǥͯ��m���", "student", _editor.Student.ID, sb.ToString());
                }

                FISCA.Presentation.Controls.MsgBox.Show("�ק�ǥͯ��m��Ʀ��\!");
                #endregion
            }
            this.Close();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #region ���������B�z

        private void dataGridViewX1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dataGridViewX1.EndEdit();
        }

        private void dataGridViewX1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dataGridViewX1.EndEdit();
        }

        private void dataGridViewX1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
                return;

            foreach (RadioButton var in flpAbsence.Controls)
            {
                if (var.Checked)
                {
                    _checkedAbsence = (K12.Data.AbsenceMappingInfo)var.Tag;
                    DataGridViewCell dgvCell = dataGridViewX1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    //�p�G����,�N�M��
                    if ("" + dgvCell.Value == _checkedAbsence.Abbreviation)
                    {
                        dgvCell.Value = "";
                        return;
                    }
                    dgvCell.Value = _checkedAbsence.Abbreviation;
                }
            }
        }

        private void dataGridViewX1_KeyUp(object sender, KeyEventArgs e)
        {
            string CurrentCellValue = ("" + dataGridViewX1.CurrentCell.Value).Trim();

            if (CheckKeys(e) && "" + CurrentCellValue == "")
            {
                return;
            }


            //�p�G����J���(�p�G�O���ť���)
            if (string.IsNullOrEmpty(CurrentCellValue))
            {
                foreach (DataGridViewCell cell in dataGridViewX1.SelectedCells)
                {
                    cell.Value = ""; //�M�Ÿ��
                }
                dataGridViewX1.GoToNEXTCell();
            }
            else
            {
                //�p�G���O�ťո��,�N����Ӫ�
                string Abbreviation = "";
                foreach (K12.Data.AbsenceMappingInfo absenceInfo in absenceList)
                {
                    if (absenceInfo.HotKey.ToUpper() == CurrentCellValue.ToUpper())
                    {
                        Abbreviation = absenceInfo.Abbreviation;
                        break;
                    }
                }

                if (Abbreviation != "") //�����,�N�������Y�g
                {
                    foreach (DataGridViewCell cell in dataGridViewX1.SelectedCells)
                    {
                        cell.Value = Abbreviation;
                    }
                }
                else //�S��������,�N�M�Ÿ��
                {
                    foreach (DataGridViewCell cell in dataGridViewX1.SelectedCells)
                    {
                        cell.Value = "";
                    }
                }
                dataGridViewX1.GoToNEXTCell();
            }
        }

        private bool CheckKeys(KeyEventArgs e)
        {
            if (e.Alt || e.Control || e.Shift ||
                e.KeyCode == Keys.Up ||
                e.KeyCode == Keys.Down ||
                e.KeyCode == Keys.Left ||
                e.KeyCode == Keys.Right ||
                e.KeyCode == Keys.Tab ||
                e.KeyCode == Keys.ShiftKey ||
                e.KeyCode == (Keys.ShiftKey | Keys.LButton) ||
                e.KeyCode == (Keys.ShiftKey | Keys.RButton) ||
                e.KeyCode == (Keys.ControlKey | Keys.LButton) ||
                e.KeyCode == (Keys.ControlKey | Keys.RButton))
            {
                return true;
            }
            else
            {
                return false;
            }

            //if (e.KeyValue < 123 && e.KeyValue > 64)
        }

        #endregion

        private void dataGridViewX1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                row.Cells[e.ColumnIndex].Selected = true;
            }
        }

        #region ���ĪG�O����DataGridView�@�}�l�|����@�檺�ĪG

        bool SelectedFalse = true;

        private void dataGridViewX1_SelectionChanged(object sender, EventArgs e)
        {
            if (SelectedFalse)
            {
                dataGridViewX1.Rows[0].Cells[0].Selected = false;
                SelectedFalse = false;
            }
        }
        #endregion
    }
}