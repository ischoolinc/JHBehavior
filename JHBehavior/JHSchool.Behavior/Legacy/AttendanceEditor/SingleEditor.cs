using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
//using SmartSchool.Common;
using FISCA.DSAUtil;
using System.Xml;
using DevComponents.DotNetBar;
using JHSchool.Behavior.Properties;
using Framework.Feature;
using Framework;
using Framework.Legacy;
using FISCA.LogAgent;
using JHSchool.Behavior.Legacy.AttendanceEditor;
using System.Xml.Linq;
using System.Xml.XPath;
//using SmartSchool.StudentRelated;
//using SmartSchool.ApplicationLog;
//using SmartSchool.Properties;

namespace JHSchool.Behavior.Legacy
{
    public partial class SingleEditor : FISCA.Presentation.Controls.BaseForm
    {
        private AbsenceInfo _checkedAbsence;
        private Dictionary<string, AbsenceInfo> _absenceList;//���O�M��
        private StudentRecord _student;
        private ISemester _semesterProvider;
        private int _startIndex;
        private ErrorProvider _errorProvider;
        private DateTime _currentStartDate;
        private DateTime _currentEndDate;
        private DateTime? _occurDate=null;

        Dictionary<string, int> ColumnIndex = new Dictionary<string, int>();

        List<string> WeekDay = new List<string>();
        List<DayOfWeek> nowWeekDay = new List<DayOfWeek>();
        List<DateTime> _Holidays = new List<DateTime>();

        //System.Windows.Forms.ToolTip toolTip = new System.Windows.Forms.ToolTip();

        //log �ݭn�Ψ쪺
        private Dictionary<string, Dictionary<string, string>> beforeData = new Dictionary<string, Dictionary<string, string>>();
        private Dictionary<string, Dictionary<string, string>> afterData = new Dictionary<string, Dictionary<string, string>>();
        private List<string> deleteData = new List<string>();

        public SingleEditor(StudentRecord student)
        {
            InitializeComponent(); //�]�p�u�㲣�ͪ�

            _errorProvider = new ErrorProvider();
            _student = student;
            _absenceList = new Dictionary<string, AbsenceInfo>();
            _semesterProvider = SemesterProvider.GetInstance();
            //_startIndex = 4;

        }

        // �ǥͤ���ί��m�n��
        public SingleEditor(StudentRecord student,DateTime occurDate)
        {
            InitializeComponent(); //�]�p�u�㲣�ͪ�

            _errorProvider = new ErrorProvider();
            _student = student;
            _absenceList = new Dictionary<string, AbsenceInfo>();
            _semesterProvider = SemesterProvider.GetInstance();
            this._occurDate = occurDate;
            //_startIndex = 4;

        }

        private void SingleEditor_Load(object sender, EventArgs e)
        {
            #region Load
            this.Text = "��H�h�ѯ��m�޲z";
            StringBuilder sb = new StringBuilder();
            string ClassName = _student.Class != null ? _student.Class.Name : "";
            sb.Append("�Z�šG<b>" + ClassName + "</b>�@");
            sb.Append("�y���G<b>" + _student.SeatNo + "</b>�@");
            sb.Append("�m�W�G<b>" + _student.Name + "</b>�@");
            sb.Append("�Ǹ��G<b>" + _student.StudentNumber + "</b>");
            lblInfo.Text = sb.ToString();
            InitializeRadioButton();
            InitializeDateRange(); //���o����w�q
            InitializeDataGridViewColumn();
            //SearchDateRange();
            //GetAbsense();
            LoadAbsense();

            #endregion
        }

        private void InitializeDateRange()
        {
            if (this._occurDate != null)
            {
                dateTimeInput1.Value =(DateTime) this._occurDate;
                dateTimeInput2.Value =(DateTime) this._occurDate;
            }
            else
            {
                #region ����w�q
                K12.Data.Configuration.ConfigData DateConfig = K12.Data.School.Configuration["Attendance_BatchEditor"];

                string date = DateConfig["SingleEditor"];

                if (date == "")
                {
                    DSXmlHelper helper = new DSXmlHelper("xml");
                    helper.AddElement("StartDate");
                    helper.AddText("StartDate", DateTime.Today.AddDays(-6).ToShortDateString());
                    helper.AddElement("EndDate");
                    helper.AddText("EndDate", DateTime.Today.ToShortDateString());
                    helper.AddElement("Locked");
                    helper.AddText("Locked", "false");

                    date = helper.BaseElement.OuterXml;
                    DateConfig["SingleEditor"] = date;
                    DateConfig.Save(); //�x�s���w�]��
                }

                XmlElement loadXml = DSXmlHelper.LoadXml(date);
                checkBoxX1.Checked = bool.Parse(loadXml.SelectSingleNode("Locked").InnerText);

                if (checkBoxX1.Checked) //�p�G�O��w,�N����w���
                {
                    dateTimeInput1.Text = loadXml.SelectSingleNode("StartDate").InnerText;
                    dateTimeInput2.Text = loadXml.SelectSingleNode("EndDate").InnerText;
                }
                else //�p�G�S����w,�N�����
                {
                    dateTimeInput1.Text = DateTime.Today.AddDays(-6).ToShortDateString();
                    dateTimeInput2.Text = DateTime.Today.ToShortDateString();
                }
                _currentStartDate = dateTimeInput1.Value;
                _currentEndDate = dateTimeInput2.Value;
                #endregion
            }
        }

        private void SaveDateSetting()
        {
            #region �x�s������
            K12.Data.Configuration.ConfigData DateConfig = K12.Data.School.Configuration["Attendance_BatchEditor"];

            DSXmlHelper helper = new DSXmlHelper("xml");
            helper.AddElement("StartDate");
            helper.AddText("StartDate", dateTimeInput1.Value.ToShortDateString());
            helper.AddElement("EndDate");
            helper.AddText("EndDate", dateTimeInput2.Value.ToShortDateString());
            helper.AddElement("Locked");
            helper.AddText("Locked", checkBoxX1.Checked.ToString());

            DateConfig["SingleEditor"] = helper.BaseElement.OuterXml;
            DateConfig.Save(); //�x�s���w�]��

            #endregion
        }

        private void InitializeRadioButton()
        {
            #region ���m���O�إ�
            DSResponse dsrsp = Config.GetAbsenceList();
            DSXmlHelper helper = dsrsp.GetContent();
            foreach (XmlElement element in helper.GetElements("Absence"))
            {
                AbsenceInfo info = new AbsenceInfo(element);
                //���䤣����
                if (!_absenceList.ContainsKey(info.Hotkey.ToUpper()))
                {
                    _absenceList.Add(info.Hotkey.ToUpper(), info);
                }
                else
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append("���m�O�G{0}\n����G{1} �w����\n(�^��r���j�p�g�����ۦP����)");
                    MsgBox.Show(string.Format(sb.ToString(), info.Name, info.Hotkey));
                }

                RadioButton rb = new RadioButton();
                rb.Text = info.Name + "(" + info.Hotkey + ")";
                rb.AutoSize = true;
                rb.Font = new Font(Framework.DotNetBar.FontStyles.GeneralFontFamily, 9.25f);
                rb.Tag = info;
                rb.CheckedChanged += new EventHandler(rb_CheckedChanged);
                rb.Click += new EventHandler(rb_CheckedChanged);
                panel.Controls.Add(rb);
                if (_checkedAbsence == null)
                {
                    _checkedAbsence = info;
                    rb.Checked = true;
                }
            }
            #endregion
        }

        void rb_CheckedChanged(object sender, EventArgs e)
        {
            #region ���m���O�إ�(�ƥ�)
            RadioButton rb = sender as RadioButton;
            if (rb.Checked) 
            {
                _checkedAbsence = rb.Tag as AbsenceInfo;
                foreach (DataGridViewCell cell in dataGridView.SelectedCells)
                {
                    if (cell.ColumnIndex < _startIndex || cell.OwningRow.Visible == false) continue;
                    cell.Value = _checkedAbsence.Abbreviation;
                    AbsenceCellInfo acInfo = cell.Tag as AbsenceCellInfo;
                    if (acInfo == null)
                    {
                        acInfo = new AbsenceCellInfo();
                    }
                    acInfo.SetValue(_checkedAbsence);
                    cell.Value = acInfo.AbsenceInfo.Abbreviation;
                    cell.Tag = acInfo;
                }
                dataGridView.Focus();
            }
            #endregion
        }
        //��l��DataGri
        private void InitializeDataGridViewColumn()
        {
            #region DataGridView��Column�إ�

            ColumnIndex.Clear(); //�M��

            DSResponse dsrsp = Config.GetPeriodList();
            DSXmlHelper helper = dsrsp.GetContent();
            //��ꧨ�O�M�ĴX�`���
            PeriodCollection collection = new PeriodCollection();
            foreach (XmlElement element in helper.GetElements("Period"))
            {
                //���o���O�θ`��
                PeriodInfo info = new PeriodInfo(element);
                collection.Items.Add(info);
            }
            //�ʺA����DataGridView 
            int ColumnsIndex = dataGridView.Columns.Add("colDate", "���");
            ColumnIndex.Add("���", ColumnsIndex);
            //dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;
            dataGridView.Columns[ColumnsIndex].Width = 120;

            ColumnsIndex = dataGridView.Columns.Add("colWeek", "�P��");
            ColumnIndex.Add("�P��", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;

            ColumnsIndex = dataGridView.Columns.Add("colSchoolYear", "�Ǧ~��");
            ColumnIndex.Add("�Ǧ~��", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = false;

            ColumnsIndex = dataGridView.Columns.Add("colSemester", "�Ǵ�");
            ColumnIndex.Add("�Ǵ�", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = false;
            dataGridView.Columns[ColumnsIndex].Frozen = true;

            _startIndex = ColumnIndex["�Ǵ�"] + 1;
            //���Ư��mDetail��JDataGridView
            foreach (PeriodInfo info in collection.GetSortedList())
            {
                int columnIndex = dataGridView.Columns.Add(info.Name, info.Name);
                DataGridViewColumn column = dataGridView.Columns[columnIndex];
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
                column.ReadOnly = true;
                column.Tag = info;
            }
            #endregion
        }

        private void SearchDateRange()
        {
            #region ������
            DateTime start = dateTimeInput1.Value;
            DateTime end = dateTimeInput2.Value;

            DateTime date = start;
            dataGridView.Rows.Clear();

            TimeSpan ts = dateTimeInput2.Value - dateTimeInput1.Value;
            if (ts.Days > 1500)
            {
                FISCA.Presentation.Controls.MsgBox.Show("�z����F" + ts.Days.ToString() + "��\n�ѩ�������϶��L��,�Э��s�]�w����I");
                _currentStartDate = dateTimeInput1.Value = DateTime.Today;
                _currentEndDate = dateTimeInput2.Value = DateTime.Today;
                return;
            }

            while (date.CompareTo(end) <= 0)
            {
                string dateValue = date.ToShortDateString();

                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridView);
                RowTag tag = new RowTag();
                tag.Date = date;
                tag.IsNew = true;
                row.Tag = tag;

                row.Cells[ColumnIndex["���"]].Value = dateValue;
                row.Cells[ColumnIndex["�P��"]].Value = GetDayOfWeekInChinese(date.DayOfWeek);
                _semesterProvider.SetDate(date);
                row.Cells[ColumnIndex["�Ǧ~��"]].Value = _semesterProvider.SchoolYear;
                row.Cells[ColumnIndex["�Ǵ�"]].Value = _semesterProvider.Semester;
                date = date.AddDays(1);

                dataGridView.Rows.Add(row);
            }
            #endregion
        }

        private void GetAbsense()
        {
            #region ���o���m�O��
            //�o�exml request
            DSXmlHelper helper = new DSXmlHelper("Request");
            helper.AddElement("Field");
            helper.AddElement("Field", "All");
            helper.AddElement("Condition");
            helper.AddElement("Condition", "RefStudentID", _student.ID);
            helper.AddElement("Condition", "StartDate", dateTimeInput1.Value.ToShortDateString());
            helper.AddElement("Condition", "EndDate", dateTimeInput2.Value.ToShortDateString());
            DSResponse dsrsp = JHSchool.Feature.Legacy.QueryAttendance.GetAttendance(new DSRequest(helper));
            helper = dsrsp.GetContent();

            //log �M�� beforeData
            beforeData.Clear();

            foreach (XmlElement element in helper.GetElements("Attendance"))
            {
                // �o�̭n���@�ǨƱ�  �Ҧp���F���i�h
                string occurDate = element.SelectSingleNode("OccurDate").InnerText;
                string schoolYear = element.SelectSingleNode("SchoolYear").InnerText;
                string semester = element.SelectSingleNode("Semester").InnerText;
                string id = element.GetAttribute("ID");
                XmlNode dNode = element.SelectSingleNode("Detail").FirstChild;

                //log �����ק�e����� �������
                DateTime logDate;
                //log dic key�Ȭ����m���
                if (DateTime.TryParse(occurDate, out logDate))
                {
                    //�p�G�٨S���]�t�������key�N�[�J
                    if (!beforeData.ContainsKey(logDate.ToShortDateString()))
                        beforeData.Add(logDate.ToShortDateString(), new Dictionary<string, string>());
                }

                DataGridViewRow row = null;
                foreach (DataGridViewRow r in dataGridView.Rows)
                {
                    DateTime date;
                    //�� DataGridView    
                    RowTag rt = r.Tag as RowTag;

                    if (!DateTime.TryParse(occurDate, out date)) continue;
                    if (date.CompareTo(rt.Date) != 0) continue;
                    row = r;
                }

                if (row == null) continue;
                RowTag rowTag = row.Tag as RowTag;
                rowTag.IsNew = false;
                rowTag.Key = id;

                row.Cells[ColumnIndex["�Ǧ~��"]].Value = schoolYear;
                row.Cells[ColumnIndex["�Ǧ~��"]].Tag = new SemesterCellInfo(schoolYear);

                row.Cells[ColumnIndex["�Ǵ�"]].Value = semester;
                row.Cells[ColumnIndex["�Ǵ�"]].Tag = new SemesterCellInfo(semester);
                //���^dataGridView
                for (int i = _startIndex; i < dataGridView.Columns.Count; i++)
                {
                    DataGridViewColumn column = dataGridView.Columns[i];
                    PeriodInfo info = column.Tag as PeriodInfo;

                    foreach (XmlNode node in dNode.SelectNodes("Period"))
                    {
                        if (node.InnerText != info.Name) continue;
                        if (node.SelectSingleNode("@AbsenceType") == null) continue;

                        DataGridViewCell cell = row.Cells[i];
                        foreach (AbsenceInfo ai in _absenceList.Values)
                        {
                            if (ai.Name != node.SelectSingleNode("@AbsenceType").InnerText) continue;
                            AbsenceInfo ainfo = ai.Clone();
                            cell.Tag = new AbsenceCellInfo(ainfo);
                            cell.Value = ai.Abbreviation;

                            //log �����ק�e����� ���m���ӳ���
                            if (!beforeData[logDate.ToShortDateString()].ContainsKey(info.Name))
                                beforeData[logDate.ToShortDateString()].Add(info.Name, ai.Name);

                            break;
                        }
                    }
                }
            }
            #endregion
        }

        //�x�s
        private void btnSave_Click(object sender, EventArgs e)
        {
            #region Save
            if (!IsValid())
            {
                FISCA.Presentation.Controls.MsgBox.Show("������ҥ��ѡA�Эץ���A���x�s", "���ҥ���", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            DSXmlHelper InsertHelper = new DSXmlHelper("InsertRequest");
            DSXmlHelper updateHelper = new DSXmlHelper("UpdateRequest");
            List<string> deleteList = new List<string>();

            List<string> synclist = new List<string>();

            ISemester semester = SemesterProvider.GetInstance();
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                RowTag tag = row.Tag as RowTag;
                semester.SetDate(tag.Date);

                //log �����ק�᪺��� �������
                if (!afterData.ContainsKey(tag.Date.ToShortDateString()))
                    afterData.Add(tag.Date.ToShortDateString(), new Dictionary<string, string>());

                if (tag.IsNew)
                {
                    #region IsNew
                    DSXmlHelper h2 = new DSXmlHelper("Attendance");
                    bool hasContent = false;
                    for (int i = _startIndex; i < dataGridView.Columns.Count; i++)
                    {
                        DataGridViewCell cell = row.Cells[i];
                        if (string.IsNullOrEmpty(("" + cell.Value).Trim())) continue;

                        PeriodInfo pinfo = dataGridView.Columns[i].Tag as PeriodInfo;
                        AbsenceCellInfo acInfo = cell.Tag as AbsenceCellInfo;
                        AbsenceInfo ainfo = acInfo.AbsenceInfo;
                        XmlElement element = h2.AddElement("Period");
                        element.InnerText = pinfo.Name;
                        element.SetAttribute("AbsenceType", ainfo.Name);
                        element.SetAttribute("AttendanceType", pinfo.Type);
                        hasContent = true;

                        //log �����ק�᪺��� ���m���ӳ���
                        if (!afterData[tag.Date.ToShortDateString()].ContainsKey(pinfo.Name))
                            afterData[tag.Date.ToShortDateString()].Add(pinfo.Name, ainfo.Name);

                    }
                    if (hasContent)
                    {
                        InsertHelper.AddElement("Attendance");
                        InsertHelper.AddElement("Attendance", "Field");
                        InsertHelper.AddElement("Attendance/Field", "RefStudentID", _student.ID);
                        InsertHelper.AddElement("Attendance/Field", "SchoolYear", row.Cells[ColumnIndex["�Ǧ~��"]].Value.ToString());
                        InsertHelper.AddElement("Attendance/Field", "Semester", row.Cells[ColumnIndex["�Ǵ�"]].Value.ToString());
                        InsertHelper.AddElement("Attendance/Field", "OccurDate", tag.Date.ToShortDateString());
                        InsertHelper.AddElement("Attendance/Field", "Detail", h2.GetRawXml(), true);
                    }

                    #endregion
                }
                else // �Y�O�쥻�N��������
                {
                    #region �O�ª�
                    DSXmlHelper h2 = new DSXmlHelper("Attendance");
                    bool hasContent = false;
                    for (int i = _startIndex; i < dataGridView.Columns.Count; i++)
                    {
                        DataGridViewCell cell = row.Cells[i];
                        if (string.IsNullOrEmpty(("" + cell.Value).Trim())) continue;

                        PeriodInfo pinfo = dataGridView.Columns[i].Tag as PeriodInfo;
                        AbsenceCellInfo acInfo = cell.Tag as AbsenceCellInfo;
                        AbsenceInfo ainfo = acInfo.AbsenceInfo;

                        XmlElement element = h2.AddElement("Period");
                        element.InnerText = pinfo.Name;
                        element.SetAttribute("AbsenceType", ainfo.Name);
                        element.SetAttribute("AttendanceType", pinfo.Type);
                        hasContent = true;

                        //log �����ק�᪺��� ���m���ӳ���
                        if (!afterData[tag.Date.ToShortDateString()].ContainsKey(pinfo.Name))
                            afterData[tag.Date.ToShortDateString()].Add(pinfo.Name, ainfo.Name);
                    }

                    if (hasContent)
                    {
                        updateHelper.AddElement("Attendance");
                        updateHelper.AddElement("Attendance", "Field");
                        updateHelper.AddElement("Attendance/Field", "RefStudentID", _student.ID);
                        updateHelper.AddElement("Attendance/Field", "SchoolYear", row.Cells[ColumnIndex["�Ǧ~��"]].Value.ToString());
                        updateHelper.AddElement("Attendance/Field", "Semester", row.Cells[ColumnIndex["�Ǵ�"]].Value.ToString());
                        updateHelper.AddElement("Attendance/Field", "OccurDate", tag.Date.ToShortDateString());
                        updateHelper.AddElement("Attendance/Field", "Detail", h2.GetRawXml(), true);
                        updateHelper.AddElement("Attendance", "Condition");
                        updateHelper.AddElement("Attendance/Condition", "ID", tag.Key);
                    }
                    else
                    {
                        deleteList.Add(tag.Key);

                        //log �����Q�R�������
                        afterData.Remove(tag.Date.ToShortDateString());
                        deleteData.Add(tag.Date.ToShortDateString());
                    }
                }
                #endregion
            }

            #region InsertHelper
            if (InsertHelper.GetElements("Attendance").Length > 0)
            {
                try
                {
                    JHSchool.Feature.Legacy.EditAttendance.Insert(new DSRequest(InsertHelper));
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("���m�����s�W���� : " + ex.Message, "�s�W����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


                //log �g�Jlog
                foreach (string date in afterData.Keys)
                {
                    if (!beforeData.ContainsKey(date) && afterData[date].Count > 0)
                    {
                        StringBuilder desc = new StringBuilder("");
                        desc.AppendLine("�ǥ͡u" + Student.Instance.Items[_student.ID].Name + "�v");
                        desc.AppendLine("����u" + date + "�v");
                        foreach (string period in afterData[date].Keys)
                        {
                            desc.AppendLine("�`���u" + period + "�v�]���u" + afterData[date][period] + "�v");
                        }

                        ApplicationLog.Log("�ǰȨt��.���m���", "�妸�s�W���m���", "student", _student.ID, desc.ToString());
                        //Log����
                        //CurrentUser.Instance.AppLog.Write(EntityType.Student, EntityAction.Insert, _student.ID, desc.ToString(), this.Text, "");
                    }
                }
            }
            #endregion

            #region updateHelper
            if (updateHelper.GetElements("Attendance").Length > 0)
            {
                try
                {
                    JHSchool.Feature.Legacy.EditAttendance.Update(new DSRequest(updateHelper));
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("���m������s���� : " + ex.Message, "��s����", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //log �g�Jlog
                foreach (string date in afterData.Keys)
                {
                    if (beforeData.ContainsKey(date) && afterData[date].Count > 0)
                    {
                        bool dirty = false;
                        StringBuilder desc = new StringBuilder("");
                        desc.AppendLine("�ǥ͡u" + Student.Instance.Items[_student.ID].Name + "�v ");
                        desc.AppendLine("����u" + date + "�v ");
                        foreach (string period in beforeData[date].Keys)
                        {
                            if (!afterData[date].ContainsKey(period))
                                afterData[date].Add(period, "");
                        }
                        foreach (string period in afterData[date].Keys)
                        {
                            if (beforeData[date].ContainsKey(period))
                            {
                                if (beforeData[date][period] != afterData[date][period])
                                {
                                    dirty = true;
                                    desc.AppendLine("�`���u" + period + "�v�ѡu" + beforeData[date][period] + "�v�ܧ󬰡u" + afterData[date][period] + "�v");
                                }
                            }
                            else
                            {
                                dirty = true;
                                desc.AppendLine("�`���u" + period + "�v�ѡu�v�ܧ󬰡u" + afterData[date][period] + "�v ");
                            }

                        }
                        if (dirty)
                        {
                            //Log����
                            ApplicationLog.Log("�ǰȨt��.���m���", "�妸�ק���m���", "student", _student.ID, desc.ToString());
                            //CurrentUser.Instance.AppLog.Write(EntityType.Student, EntityAction.Update, _student.ID, desc.ToString(), this.Text, "");
                        }
                    }
                }
            }
            #endregion

            #region deleteList
            if (deleteList.Count > 0)
            {
                DSXmlHelper deleteHelper = new DSXmlHelper("DeleteRequest");
                deleteHelper.AddElement("Attendance");
                foreach (string key in deleteList)
                {
                    deleteHelper.AddElement("Attendance", "ID", key);
                }

                try
                {
                    JHSchool.Feature.Legacy.EditAttendance.Delete(new DSRequest(deleteHelper));
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("���m�����R������ : " + ex.Message, "�R������", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //log �g�J�Q�R������ƪ�log
                StringBuilder desc = new StringBuilder("");
                desc.AppendLine("�ǥ͡u" + Student.Instance.Items[_student.ID].Name + "�v");
                foreach (string date in deleteData)
                {
                    desc.AppendLine("�R���u" + date + "�v���m���� ");
                }
                //Log����
                ApplicationLog.Log("�ǰȨt��.���m���", "�妸�R�����m���", "student", _student.ID, desc.ToString());
                //CurrentUser.Instance.AppLog.Write(EntityType.Student, EntityAction.Delete, _student.ID, desc.ToString(), this.Text, "");
            }
            #endregion

            //Ĳ�o�ܧ�ƥ�
            Attendance.Instance.SyncDataBackground(_student.ID);
            //Student.Instance.SyncDataBackground(_student.ID);

            FISCA.Presentation.Controls.MsgBox.Show("�x�s���m��Ʀ��\!", "����", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
            #endregion
        }

        private bool IsValid()
        {
            #region DataGridView�������(�p�GErrorText���e����)
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.ErrorText != string.Empty)
                        return false;
                }
            }
            return true;
            #endregion
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            LoadAbsense(); //���s��z
        }

        private void LoadAbsense()
        {
            #region �B�z�P���]�w��
            WeekDay.Clear();
            nowWeekDay.Clear();
            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["���m�妸�n��_�P���]�w"];
            string cdIN = cd["�P���]�w"];

            XmlElement day;

            if (cdIN != "")
            {
                day = XmlHelper.LoadXml(cdIN);
            }
            else
            {
                day = null;
            }

            if (day != null)
            {
                foreach (XmlNode each in day.SelectNodes("Day"))
                {
                    XmlElement each2 = each as XmlElement;
                    WeekDay.Add(each2.GetAttribute("Detail"));
                }
            }
            else
            {
                WeekDay.AddRange(new string[] { "�P���@", "�P���G", "�P���T", "�P���|", "�P����", "�P����", "�P����" });
            }

            nowWeekDay = ChengDayOfWeel(WeekDay);
            #endregion

            #region ���o���O�M��
            ConfigData _CD;
            //���o���e�]�w�]�w
            _CD = School.Configuration["SCHOOL_HOLIDAY_CONFIG_STRING"];
            XElement rootXml = null;
            string xmlContent = _CD["CONFIG_STRING"];

            if (!string.IsNullOrWhiteSpace(xmlContent))
                rootXml = XElement.Parse(xmlContent);
            else
                rootXml = new XElement("SchoolHolidays");

            //Ū�X���e������M��
            foreach (XElement holiday in rootXml.XPathSelectElements("//Holiday"))
            {
                DateTime date;
                if (DateTime.TryParse(holiday.Value, out date))
                    _Holidays.Add(date);
            }
            #endregion

            SearchDateRange();
            GetAbsense();
            filterRows(null, null);

        }

        private List<DayOfWeek> ChengDayOfWeel(List<string> list)
        {
            #region ���o�P����Ӫ�
            List<DayOfWeek> DOW = new List<DayOfWeek>();
            foreach (string each in list)
            {
                if (each == "�P���@")
                {
                    DOW.Add(DayOfWeek.Monday);
                }
                else if (each == "�P���G")
                {
                    DOW.Add(DayOfWeek.Tuesday);
                }
                else if (each == "�P���T")
                {
                    DOW.Add(DayOfWeek.Wednesday);
                }
                else if (each == "�P���|")
                {
                    DOW.Add(DayOfWeek.Thursday);
                }
                else if (each == "�P����")
                {
                    DOW.Add(DayOfWeek.Friday);
                }
                else if (each == "�P����")
                {
                    DOW.Add(DayOfWeek.Saturday);
                }
                else if (each == "�P����")
                {
                    DOW.Add(DayOfWeek.Sunday);
                }
            }

            return DOW;
            #endregion
        }

        private string GetDayOfWeekInChinese(DayOfWeek day)
        {
            #region �P��(��/�^)��Ӫ�
            switch (day)
            {
                case DayOfWeek.Monday:
                    return "�@";
                case DayOfWeek.Tuesday:
                    return "�G";
                case DayOfWeek.Wednesday:
                    return "�T";
                case DayOfWeek.Thursday:
                    return "�|";
                case DayOfWeek.Friday:
                    return "��";
                case DayOfWeek.Saturday:
                    return "��";
                default:
                    return "��";
            }
            #endregion
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            if (IsDirty())
            {
                if (FISCA.Presentation.Controls.MsgBox.Show("��Ƥw�ܧ�B�|���x�s�A�O�_���w�s����?", "�T�{", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    this.Close();
                }
            }
            else
                this.Close();
        }

        private void dataGridView_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            #region CellDoubleClick
            if (e.Button != MouseButtons.Left) return;
            if (e.ColumnIndex < _startIndex) return;
            if (e.RowIndex < 0) return;
            DataGridViewCell cell = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
            cell.Value = _checkedAbsence.Abbreviation;
            AbsenceCellInfo acInfo = cell.Tag as AbsenceCellInfo;
            if (acInfo == null)
            {
                acInfo = new AbsenceCellInfo();
            }
            acInfo.SetValue(_checkedAbsence);
            if (acInfo.IsValueChanged)
                cell.Value = acInfo.AbsenceInfo.Abbreviation;
            else
            {
                cell.Value = string.Empty;
                acInfo.SetValue(AbsenceInfo.Empty);
            }
            cell.Tag = acInfo;
            #endregion
        }

        private void dataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            #region �Ǧ~��/�Ǵ���J����
            DataGridViewCell cell = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
            if (e.ColumnIndex == ColumnIndex["�Ǧ~��"])
            {
                string errorMessage = "";
                int schoolYear;
                if (cell.Value == null)
                    errorMessage = "�Ǧ~�פ��i���ť�";
                else if (!int.TryParse(cell.Value.ToString(), out schoolYear))
                    errorMessage = "�Ǧ~�ץ��������";

                if (errorMessage != "")
                {
                    cell.Style.BackColor = Color.Red;
                    cell.ToolTipText = errorMessage;
                }
                else
                {
                    cell.Style.BackColor = Color.White;
                    cell.ToolTipText = "";
                }
            }
            else if (e.ColumnIndex == ColumnIndex["�Ǵ�"])
            {
                string errorMessage = "";

                if (cell.Value == null)
                    errorMessage = "�Ǵ����i���ť�";
                else if (cell.Value.ToString() != "1" && cell.Value.ToString() != "2")
                    errorMessage = "�Ǵ���������ơy1�z�Ρy2�z";

                if (errorMessage != "")
                {
                    cell.ErrorText = errorMessage;
                }
                else
                {
                    cell.ErrorText = string.Empty;
                }
            }
            #endregion
        }

        private void dataGridView_KeyDown(object sender, KeyEventArgs e)
        {
            #region �p�G���U���s
            string key = KeyConverter.GetKeyMapping(e);

            if (!_absenceList.ContainsKey(key))
            {
                if (e.KeyCode != Keys.Space && e.KeyCode != Keys.Delete) return;
                foreach (DataGridViewCell cell in dataGridView.SelectedCells)
                {
                    if (cell.ColumnIndex < _startIndex || cell.OwningRow.Visible == false) continue;
                    cell.Value = null;
                    AbsenceCellInfo acInfo = cell.Tag as AbsenceCellInfo;
                    if (acInfo != null)
                        acInfo.SetValue(null);
                }
            }
            else
            {
                AbsenceInfo info = _absenceList[key];
                foreach (DataGridViewCell cell in dataGridView.SelectedCells)
                {
                    if (cell.ColumnIndex < _startIndex || cell.OwningRow.Visible == false) continue;
                    AbsenceCellInfo acInfo = cell.Tag as AbsenceCellInfo;

                    if (acInfo == null)
                    {
                        acInfo = new AbsenceCellInfo();
                    }
                    acInfo.SetValue(info);

                    if (acInfo.IsValueChanged)
                        cell.Value = acInfo.AbsenceInfo.Abbreviation;
                    else
                    {
                        cell.Value = string.Empty;
                        acInfo.SetValue(AbsenceInfo.Empty);
                    }
                    cell.Tag = acInfo;
                }
            }
            #endregion
        }

        //private void picLock_Click(object sender, EventArgs e)
        //{
        //    bool isLock = false;
        //    if (picLock.Tag != null)
        //    {
        //        if (!bool.TryParse(picLock.Tag.ToString(), out isLock))
        //            isLock = false;
        //    }
        //    if (isLock)
        //    {
        //        picLock.Image = Resources.unlock;
        //        picLock.Tag = false;
        //        toolTip.SetToolTip(picLock, "���m�n�����������w���A�A�z�i�H�I��ϥܡA�N�S�w����϶���w�C");
        //        labelX2.Text = "";
        //    }
        //    else
        //    {
        //        picLock.Image = Resources._lock;
        //        picLock.Tag = true;
        //        toolTip.SetToolTip(picLock, "���m�n������w��w�A�z�i�H�I��ϥܸѰ���w�C");
        //        labelX2.Text = "�w��w���m���";
        //    }
        //    SaveDateSetting();
        //}

        private void dateTimeInput1_Validated(object sender, EventArgs e)
        {
            #region dateTimeInput1����ܧ�ƥ�
            _errorProvider.SetError(dateTimeInput1, string.Empty);

            if (IsDirty())
            {
                if (FISCA.Presentation.Controls.MsgBox.Show("��Ƥw�ܧ�B�|���x�s�A�O�_���w�s����?", "�T�{", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    dateTimeInput1.Value = _currentStartDate;
                    return;
                }
            }
            _currentStartDate = dateTimeInput1.Value;
            dataGridView.Rows.Clear();
            LoadAbsense();
            #endregion
        }


        private void dateTimeInput2_Validated(object sender, EventArgs e)
        {
            #region dateTimeInput1����ܧ�ƥ�
            _errorProvider.SetError(dateTimeInput2, string.Empty);

            if (IsDirty())
            {
                if (FISCA.Presentation.Controls.MsgBox.Show("��Ƥw�ܧ�B�|���x�s�A�O�_���w�s����?", "�T�{", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    dateTimeInput2.Value = _currentEndDate;
                    return;
                }
            }
            _currentEndDate = dateTimeInput2.Value;
            dataGridView.Rows.Clear();
            LoadAbsense();
            #endregion
        }

        private bool IsDirty()
        {
            #region �������
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Tag == null) continue;
                    if (cell.Tag is SemesterCellInfo)
                    {
                        SemesterCellInfo cInfo = cell.Tag as SemesterCellInfo;
                        if (cInfo.IsDirty) return true;
                    }
                    else if (cell.Tag is AbsenceCellInfo)
                    {
                        AbsenceCellInfo cInfo = cell.Tag as AbsenceCellInfo;
                        if (cInfo.IsDirty) return true;
                    }
                }
            }
            return false;
            #endregion
        }

        private void filterRows(object sender, EventArgs e)
        {
            dataGridView.SuspendLayout();

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                row.Visible = true;

                bool hasData = false;
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.ColumnIndex < _startIndex || cell.OwningRow.Visible == false) continue;
                    if (!string.IsNullOrEmpty("" + cell.Value))
                    {
                        hasData = true;
                        break;
                    }
                }

                bool isHolday = false;
                RowTag rowTag = row.Tag as RowTag;
                if (!nowWeekDay.Contains(rowTag.Date.DayOfWeek)) //�z�藍��ܪ��P��
                {
                    isHolday = true;
                }
                else if (_Holidays.Contains(rowTag.Date)) //�z�ﰲ��
                {
                    isHolday = true;
                }

                if (hasData)
                {
                    if (isHolday)
                    {
                        row.Cells[ColumnIndex["�P��"]].Style.ForeColor = Color.Red;
                    }
                }
                else
                {
                    if (chkHasData.Checked == true)
                        row.Visible = false;
                    if (isHolday)
                        row.Visible = false;
                }
            }

            dataGridView.ResumeLayout();
        }

        private void btnDay_Click(object sender, EventArgs e)
        {
            #region �P���]�w
            Searchday Sday = new Searchday("���m�妸�n��_�P���]�w");
            if (Sday.ShowDialog() == DialogResult.Yes)
            {
                LoadAbsense();
            }
            #endregion
        }

        private void dateTimeInput1_TextChanged(object sender, EventArgs e)
        {
            if (dataGridView.Rows.Count != 0)
            {
                LoadAbsense();
            }
        }

        private void dateTimeInput2_TextChanged(object sender, EventArgs e)
        {
            if (dataGridView.Rows.Count != 0)
            {
                LoadAbsense();
            }
        }

        private void SingleEditor_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveDateSetting();
        }

        private void dataGridView_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                row.Cells[e.ColumnIndex].Selected = true;
            }
        }

    }

    class RowTag
    {
        #region RowTag

        private DateTime _date;

        public DateTime Date
        {
            get { return _date; }
            set { _date = value; }
        }
        private bool _isNew;

        public bool IsNew
        {
            get { return _isNew; }
            set { _isNew = value; }
        }

        private string _key;

        public string Key
        {
            get { return _key; }
            set { _key = value; }
        }

        #endregion
    }
}