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

        //Log使用
        //之前
        private Dictionary<string, string> DicBeforeLog = new Dictionary<string, string>();
        //之後
        private Dictionary<string, string> DicAfterLog = new Dictionary<string, string>();

        public AttendanceForm(EditorStatus status, JHAttendanceRecord editor, List<K12.Data.PeriodMappingInfo> periodList, FeatureAce permission, int SchoolYear, int Semester)
        {
            InitializeComponent();

            #region 初始化學年度學期

            //學年度
            cboSchool.Items.Add(int.Parse(JHSchool.School.DefaultSchoolYear) - 4);
            cboSchool.Items.Add(int.Parse(JHSchool.School.DefaultSchoolYear) - 3);
            cboSchool.Items.Add(int.Parse(JHSchool.School.DefaultSchoolYear) - 2);
            cboSchool.Items.Add(int.Parse(JHSchool.School.DefaultSchoolYear) - 1);
            int SchoolYearSelectIndex = cboSchool.Items.Add(SchoolYear);
            cboSchool.Items.Add(int.Parse(JHSchool.School.DefaultSchoolYear) + 1);
            cboSchool.Items.Add(int.Parse(JHSchool.School.DefaultSchoolYear) + 2);
            cboSchool.Items.Add(int.Parse(JHSchool.School.DefaultSchoolYear) + 3);
            cboSchool.SelectedIndex = SchoolYearSelectIndex;
            //學期
            cboSemester.Items.Add(1);
            cboSemester.Items.Add(2);
            cboSemester.SelectedIndex = Semester == 1 ? 0 : 1;

            #endregion

            #region 初始化缺曠類別

            //初始化時,即取得最新缺曠資料
            absenceList = K12.Data.AbsenceMapping.SelectAll();

            foreach (K12.Data.AbsenceMappingInfo info in absenceList)
            {
                RadioButton rb = new RadioButton();
                //缺曠別,縮寫,熱鍵
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

            //把第一個缺曠類別設為預設值
            RadioButton fouse = flpAbsence.Controls[0] as RadioButton;
            fouse.Checked = true;

            #endregion

            #region 初始化節次表
            foreach (K12.Data.PeriodMappingInfo info in periodList)
            {
                //Log使用
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
            cboSchool.Enabled = permission.Editable;
            cboSemester.Enabled = permission.Editable;
            panelAbsence.Enabled = permission.Editable;
            //pancelAttendence.Enabled = permission.Editable;
            dataGridViewX1.Enabled = permission.Editable;

            if (status == EditorStatus.Insert)
            {
                Text = "管理學生缺曠紀錄【新增模式】";
            }

            _status = status;
            _editor = editor;
            _absenceList = absenceList;
            _periodList = periodList;
        }

        //void TextBox_MouseDoubleClick(object sender, MouseEventArgs e)
        //{
        //    foreach (RadioButton var in flpAbsence.Controls)
        //    {
        //        if (var.Checked)
        //        {
        //            _checkedAbsence = (AbsenceMappingInfo)var.Tag;

        //            TextBox textBox = sender as TextBox;
        //            if (textBox.Text == _checkedAbsence.Abbreviation)
        //            {
        //                textBox.Text = "";
        //                textBox.Tag = null;
        //                return;
        //            }
        //            textBox.Text = _checkedAbsence.Abbreviation;
        //        }
        //    }
        //}

        public AttendanceForm(EditorStatus status, JHAttendanceRecord editor, List<K12.Data.PeriodMappingInfo> periodList, FeatureAce permission)
        {
            InitializeComponent();

            #region 初始化學年度學期

            //學年度
            cboSchool.Items.Add(int.Parse(JHSchool.School.DefaultSchoolYear) - 3);
            cboSchool.Items.Add(int.Parse(JHSchool.School.DefaultSchoolYear) - 2);
            cboSchool.Items.Add(int.Parse(JHSchool.School.DefaultSchoolYear) - 1);
            int SchoolYearSelectIndex = cboSchool.Items.Add(int.Parse(JHSchool.School.DefaultSchoolYear));
            cboSchool.Items.Add(int.Parse(JHSchool.School.DefaultSchoolYear) + 1);
            cboSchool.Items.Add(int.Parse(JHSchool.School.DefaultSchoolYear) + 2);
            cboSchool.SelectedIndex = SchoolYearSelectIndex;
            //學期
            cboSemester.Items.Add(1);
            cboSemester.Items.Add(2);
            cboSemester.SelectedIndex = JHSchool.School.DefaultSemester == "1" ? 0 : 1;

            #endregion

            #region 初始化缺曠類別

            //初始化時,即取得最新缺曠資料
            absenceList = K12.Data.AbsenceMapping.SelectAll();

            foreach (K12.Data.AbsenceMappingInfo info in absenceList)
            {
                RadioButton rb = new RadioButton();
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

            //把第一個缺曠類別設為預設值
            RadioButton fouse = flpAbsence.Controls[0] as RadioButton;
            fouse.Checked = true;

            #endregion

            #region 初始化節次表
            foreach (K12.Data.PeriodMappingInfo info in periodList)
            {
                //Log使用
                DicBeforeLog.Add(info.Name, "");
                DicAfterLog.Add(info.Name, "");

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
            cboSchool.Enabled = permission.Editable;
            cboSemester.Enabled = permission.Editable;
            panelAbsence.Enabled = permission.Editable;
            //pancelAttendence.Enabled = permission.Editable;

            if (status == EditorStatus.Insert)
            {
                Text = "管理學生缺曠紀錄【新增模式】";
            }

            if (status == EditorStatus.Update)
            {
                Text = "管理學生缺曠紀錄【修改模式】";

                dateTimeInput1.Value = editor.OccurDate;
                dateTimeInput1.Enabled = false;
                cboSchool.Text = editor.SchoolYear.ToString();
                cboSemester.Text = editor.Semester.ToString();

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
                Text = "管理學生缺曠紀錄【檢視模式】";

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
            if (_status == EditorStatus.Insert)
            {
                foreach (JHAttendanceRecord each in JHAttendance.SelectByStudentIDs(new string[] { _editor.RefStudentID }))
                {
                    if (each.OccurDate == dateTimeInput1.Value)
                    {
                        _errorProvider.SetError(dateTimeInput1, "此日期已有紀錄存在，請改用修改模式");
                        return;
                    }
                }
            }

            _editor.OccurDate = dateTimeInput1.Value;
            _editor.SchoolYear = int.Parse(cboSchool.Text);
            _editor.Semester = int.Parse(cboSemester.Text);

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
                            //ap.AttendanceType = "一般";

                        }
                    }

                    periodDetail.Add(ap);
                }
            }

            _editor.PeriodDetail = periodDetail;

            if (periodDetail.Count == 0)
            {
                FISCA.Presentation.Controls.MsgBox.Show("不可儲存無缺曠內容之資料!");
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
                    FISCA.Presentation.Controls.MsgBox.Show("新增缺曠資料失敗!" + ex.Message);
                    return;
                }

                StringBuilder sb = new StringBuilder();
                sb.Append("學生「" + _editor.Student.Name + "」");
                sb.Append("日期「" + _editor.OccurDate.ToShortDateString() + "」");
                sb.AppendLine("新增一筆缺曠資料。");
                sb.AppendLine("詳細資料：");
                foreach (K12.Data.AttendancePeriod each in _editor.PeriodDetail)
                {
                    sb.AppendLine("節次「" + each.Period + "」設為「" + each.AbsenceType + "」");
                }

                ApplicationLog.Log("學務系統.缺曠資料", "新增學生缺曠資料", "student", _editor.Student.ID, sb.ToString());
                FISCA.Presentation.Controls.MsgBox.Show("新增學生缺曠資料成功!"); 
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
                    FISCA.Presentation.Controls.MsgBox.Show("修改缺曠資料失敗!" + ex.Message);
                    return;
                }

                StringBuilder sb = new StringBuilder();
                sb.Append("學生「" + _editor.Student.Name + "」");
                sb.Append("日期「" + _editor.OccurDate.ToShortDateString() + "」");
                sb.AppendLine("缺曠資料已修改。");
                sb.AppendLine("詳細資料：");

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
                    if (DicAfterLog[each.Name] != "" && DicBeforeLog[each.Name] != "" && DicAfterLog[each.Name] != DicBeforeLog[each.Name])
                    {
                        sb.AppendLine("節次「" + each.Name + "」由「" + DicBeforeLog[each.Name] + "」變更為「" + DicAfterLog[each.Name] + "」");
                        LogMode = true;
                    }
                }

                if (LogMode)
                {
                    ApplicationLog.Log("學務系統.缺曠資料", "修改學生缺曠資料", "student", _editor.Student.ID, sb.ToString());
                }

                FISCA.Presentation.Controls.MsgBox.Show("修改學生缺曠資料成功!"); 
                #endregion
            }
            this.Close();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #region 未完成的處理
        private void dataGridViewX1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //DataGridViewCell dgvCell = dataGridViewX1.Rows[e.RowIndex].Cells[e.ColumnIndex];
            //string dgvString = "" + dgvCell.Value;
            //dgvCell.Value = "";
            //if (!string.IsNullOrEmpty(dgvString))
            //{
            //    foreach (AbsenceMappingInfo absenceInfo in absenceList)
            //    {
            //        if (dgvString == absenceInfo.HotKey)
            //        {
            //            dgvCell.Value = absenceInfo.Abbreviation;
            //            break;
            //        }
            //    }
            //    dataGridViewX1.GoToNEXTCell();
            //}
        }

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
                    //如果有值,就清空
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
            

            //如果未輸入資料(如果是按空白鍵)
            if (string.IsNullOrEmpty(CurrentCellValue))
            {
                foreach (DataGridViewCell cell in dataGridViewX1.SelectedCells)
                {
                    cell.Value = ""; //清空資料
                }
                dataGridViewX1.GoToNEXTCell();
            }
            else
            {
                //如果不是空白資料,就比對對照表
                string Abbreviation = "";
                foreach (K12.Data.AbsenceMappingInfo absenceInfo in absenceList)
                {
                    if (absenceInfo.HotKey.ToUpper() == CurrentCellValue.ToUpper())
                    {
                        Abbreviation = absenceInfo.Abbreviation;
                        break;
                    }
                }

                if (Abbreviation != "") //對到資料,就替換為縮寫
                {
                    foreach (DataGridViewCell cell in dataGridViewX1.SelectedCells)
                    {
                        cell.Value = Abbreviation;
                    }
                }
                else //沒有比對到資料,就清空資料
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
                e.KeyCode == (Keys.ControlKey | Keys.LButton)||
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

        #region 此效果是取消DataGridView一開始會選取一格的效果

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