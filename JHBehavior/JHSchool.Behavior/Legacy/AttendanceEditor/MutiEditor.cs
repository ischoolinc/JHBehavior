using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
//using SmartSchool.Common;
//using SmartSchool.StudentRelated;
using FISCA.DSAUtil;
using System.Xml;
using DevComponents.DotNetBar;
using JHSchool.Behavior.Properties;
using JHSchool.Feature.Legacy;
using Framework;
using Framework.Legacy;
using JHSchool.Data;
using FISCA.LogAgent;
//using SmartSchool.ApplicationLog;
//using SmartSchool.Properties;

//using SmartReport.Properties;
//using SmartSchool.SmartPlugIn.Properties;

namespace JHSchool.Behavior.Legacy
{
    public partial class MutiEditor : FISCA.Presentation.Controls.BaseForm
    {
        private List<K12.Data.StudentRecord> _students;
        private ISemester _semesterProvider;
        private Dictionary<string, AbsenceInfo> _absenceList;
        private int _startIndex;
        private AbsenceInfo _checkedAbsence;
        private DateTime _currentDate;
        List<DataGridViewRow> _hiddenRows;

        Dictionary<string, int> ColumnIndex = new Dictionary<string, int>();

        //System.Windows.Forms.ToolTip toolTip = new System.Windows.Forms.ToolTip();

        //log 需要用到的
        private Dictionary<string, Dictionary<string, string>> beforeData = new Dictionary<string, Dictionary<string, string>>();
        private Dictionary<string, Dictionary<string, string>> afterData = new Dictionary<string, Dictionary<string, string>>();
        private List<string> deleteData = new List<string>();
        private DateTime logDate;

        public MutiEditor(List<K12.Data.StudentRecord> students)
        {
            InitializeComponent(); //設計工具產生的

            _students = students;
            _absenceList = new Dictionary<string, AbsenceInfo>();
            _semesterProvider = SemesterProvider.GetInstance();

            _hiddenRows = new List<DataGridViewRow>();
        }

        private void MutiEditor_Load(object sender, EventArgs e)
        {
            #region Load
            InitializeRadioButton();
            InitializeDateRange();
            InitializeDataGridView();
            //SearchStudentRange();
            btnRenew_Click(null, null);

            #region toolTip
            //toolTip.AutoPopDelay = 5000;
            //toolTip.InitialDelay = 1000;
            //toolTip.ReshowDelay = 500;
            //toolTip.ShowAlways = true;

            //bool isLock = false;
            //if (picLock.Tag != null)
            //{
            //    if (!bool.TryParse(picLock.Tag.ToString(), out isLock))
            //        isLock = false;
            //}
            //if (isLock)
            //{
            //    labelX2.Text = "已鎖定缺曠日期";
            //    toolTip.SetToolTip(picLock, "缺曠日期已鎖定，您可以點選圖示解除鎖定。");
            //}
            //else
            //{
            //    labelX2.Text = "";
            //    toolTip.SetToolTip(picLock, "缺曠日期為未鎖定狀態，您可以點選圖示，將特定日期鎖定。");
            //} 
            #endregion 
            #endregion
        }

        private void InitializeDateRange()
        {
            #region 日期定義
            K12.Data.Configuration.ConfigData DateConfig = K12.Data.School.Configuration["Attendance_BatchEditor"];

            string date = DateConfig["MutiEditor"];

            if (date == "")
            {
                DSXmlHelper helper = new DSXmlHelper("xml");
                helper.AddElement("Date");
                helper.AddText("Date", DateTime.Today.ToShortDateString());
                helper.AddElement("Locked");
                helper.AddText("Locked", "false");

                date = helper.BaseElement.OuterXml;
                DateConfig["MutiEditor"] = date;
                DateConfig.Save(); //儲存此預設檔
            }

            XmlElement loadXml = DSXmlHelper.LoadXml(date);
            checkBoxX1.Checked = bool.Parse(loadXml.SelectSingleNode("Locked").InnerText);

            if (checkBoxX1.Checked) //如果是鎖定,就取鎖定日期
            {
                dateTimeInput1.Text = loadXml.SelectSingleNode("Date").InnerText;
            }
            else //如果沒有鎖定,就取當天
            {
                dateTimeInput1.Text = DateTime.Today.ToShortDateString();
            }
            _currentDate = dateTimeInput1.Value;
            #endregion
        }

        private void SaveDateSetting()
        {
            #region 儲存日期資料
            K12.Data.Configuration.ConfigData DateConfig = K12.Data.School.Configuration["Attendance_BatchEditor"];

            DSXmlHelper helper = new DSXmlHelper("xml");
            helper.AddElement("Date");
            helper.AddText("Date", dateTimeInput1.Value.ToShortDateString());
            helper.AddElement("Locked");
            helper.AddText("Locked", checkBoxX1.Checked.ToString());

            DateConfig["MutiEditor"] = helper.BaseElement.OuterXml;
            DateConfig.Save(); //儲存此預設檔

            #endregion
        }

        private void InitializeRadioButton()
        {
            #region 缺曠類別建立
            DSResponse dsrsp = Framework.Feature.Config.GetAbsenceList();
            DSXmlHelper helper = dsrsp.GetContent();
            foreach (XmlElement element in helper.GetElements("Absence"))
            {
                AbsenceInfo info = new AbsenceInfo(element);
                //熱鍵不重覆
                if (!_absenceList.ContainsKey(info.Hotkey.ToUpper()))
                {
                    _absenceList.Add(info.Hotkey.ToUpper(), info);
                }
                else
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append("缺曠別：{0}\n熱鍵：{1} 已重覆\n(英文字母大小寫視為相同熱鍵)");
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
            #region 缺曠類別建立(事件)
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

        private void InitializeDataGridView()
        {
            InitializeDataGridViewColumn();
        }

        private void InitializeDataGridViewColumn()
        {
            #region DataGridView的Column建立

            ColumnIndex.Clear();

            DSResponse dsrsp = Framework.Feature.Config.GetPeriodList();
            DSXmlHelper helper = dsrsp.GetContent();
            PeriodCollection collection = new PeriodCollection();
            foreach (XmlElement element in helper.GetElements("Period"))
            {
                PeriodInfo info = new PeriodInfo(element);
                collection.Items.Add(info);
            }
            int ColumnsIndex = dataGridView.Columns.Add("colClassName", "班級");
            ColumnIndex.Add("班級", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;


            ColumnsIndex = dataGridView.Columns.Add("colSeatNo", "座號");
            ColumnIndex.Add("座號", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;

            ColumnsIndex = dataGridView.Columns.Add("colName", "姓名");
            ColumnIndex.Add("姓名", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;
            dataGridView.Columns[ColumnsIndex].Frozen = true; //由此開始可拖移

            ColumnsIndex = dataGridView.Columns.Add("colSchoolNumber", "學號");
            ColumnIndex.Add("學號", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;

            ColumnsIndex = dataGridView.Columns.Add("colSchoolYear", "學年度");
            ColumnIndex.Add("學年度", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = false;

            ColumnsIndex = dataGridView.Columns.Add("colSemester", "學期");
            ColumnIndex.Add("學期", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = false;

            _startIndex = ColumnIndex["學期"] + 1;

            List<string> cols = new List<string>() { "學年度", "學期" };

            foreach (PeriodInfo info in collection.GetSortedList())
            {
                cols.Add(info.Name);

                int columnIndex = dataGridView.Columns.Add(info.Name, info.Name);
                ColumnIndex.Add(info.Name, columnIndex); //節次
                DataGridViewColumn column = dataGridView.Columns[columnIndex];
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
                column.ReadOnly = true;
                column.Tag = info;
            }

            Campus.Windows.DataGridViewImeDecorator dec = new Campus.Windows.DataGridViewImeDecorator(this.dataGridView, cols);

            #endregion
        }

        public int SortStudent(K12.Data.StudentRecord x, K12.Data.StudentRecord y)
        {
            K12.Data.StudentRecord student1 = x;
            K12.Data.StudentRecord student2 = y;

            string ClassName1 = student1.Class != null ? student1.Class.Name : "";
            ClassName1 = ClassName1.PadLeft(5, '0');
            string ClassName2 = student2.Class != null ? student2.Class.Name : "";
            ClassName2 = ClassName2.PadLeft(5, '0');

            string Sean1 = student1.SeatNo.HasValue ? student1.SeatNo.Value.ToString() : "";
            Sean1 = Sean1.PadLeft(3, '0');
            string Sean2 = student2.SeatNo.HasValue ? student2.SeatNo.Value.ToString() : "";
            Sean2 = Sean2.PadLeft(3, '0');

            ClassName1 += Sean1;
            ClassName2 += Sean2;

            return ClassName1.CompareTo(ClassName2);
        }

        private void SearchStudentRange()
        {
            #region 日期選擇
            dataGridView.Rows.Clear();
            _semesterProvider.SetDate(dateTimeInput1.Value);
            //_students.Sort(SortStudent);
            _students = SortClassIndex.K12Data_StudentRecord(_students);
            foreach (K12.Data.StudentRecord student in _students)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridView);
                row.Cells[ColumnIndex["班級"]].Value = (student.Class != null) ? student.Class.Name : "";
                row.Cells[ColumnIndex["座號"]].Value = student.SeatNo;
                row.Cells[ColumnIndex["姓名"]].Value = student.Name;
                row.Cells[ColumnIndex["學號"]].Value = student.StudentNumber;
                row.Cells[ColumnIndex["學年度"]].Value = _semesterProvider.SchoolYear;
                row.Cells[ColumnIndex["學期"]].Value = _semesterProvider.Semester;

                row.Cells[ColumnIndex["學年度"]].Tag = new SemesterCellInfo(_semesterProvider.SchoolYear.ToString());
                row.Cells[ColumnIndex["學期"]].Tag = new SemesterCellInfo(_semesterProvider.Semester.ToString());
                StudentRowTag tag = new StudentRowTag();
                tag.Student = student;
                row.Tag = tag;

                dataGridView.Rows.Add(row);
            } 
            #endregion
        }

        private void GetAbsense()
        {            
            #region 取得缺曠記錄
            DSXmlHelper helper = new DSXmlHelper("Request");
            helper.AddElement("Field");
            helper.AddElement("Field", "All");
            helper.AddElement("Condition");
            helper.AddElement("Condition", "OccurDate", dateTimeInput1.Value.ToShortDateString());
            foreach (K12.Data.StudentRecord student in _students)
                helper.AddElement("Condition", "RefStudentID", student.ID);

            DSResponse dsrsp = QueryAttendance.GetAttendance(new DSRequest(helper));
            helper = dsrsp.GetContent();

            //log 清除 beforeData
            beforeData.Clear();

            //log 紀錄日期
            logDate = dateTimeInput1.Value;

            foreach (XmlElement element in helper.GetElements("Attendance"))
            {
                // 這裡要做一些事情  例如找到東西塞進去
                string occurDate = element.SelectSingleNode("OccurDate").InnerText;
                string schoolYear = element.SelectSingleNode("SchoolYear").InnerText;
                string semester = element.SelectSingleNode("Semester").InnerText;
                string studentid = element.SelectSingleNode("RefStudentID").InnerText;
                string id = element.GetAttribute("ID");
                XmlNode dNode = element.SelectSingleNode("Detail").FirstChild;

                //log 紀錄修改前的資料 紀錄學生ID
                if (!beforeData.ContainsKey(studentid))
                    beforeData.Add(studentid, new Dictionary<string, string>());

                DataGridViewRow row = null;
                foreach (DataGridViewRow r in dataGridView.Rows)
                {
                    StudentRowTag rt = r.Tag as StudentRowTag;
                    if (rt.Student.ID == studentid)
                    {
                        row = r;
                        rt.RowID = id;
                        break;
                    }
                }

                if (row == null) continue;

                row.Cells[ColumnIndex["學年度"]].Value = schoolYear;
                row.Cells[ColumnIndex["學年度"]].Tag = new SemesterCellInfo(schoolYear);

                row.Cells[ColumnIndex["學期"]].Value = semester;
                row.Cells[ColumnIndex["學期"]].Tag = new SemesterCellInfo(semester);

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

                            //log 紀錄修改前的資料 缺曠明細部分
                            if (!beforeData[studentid].ContainsKey(info.Name))
                                beforeData[studentid].Add(info.Name, ai.Name);

                            break;
                        }
                    }
                }
            } 
            #endregion
        }

        private void buttonX3_Click(object sender, EventArgs e)
        {
            #region Save
            if (!IsValid())
            {
                FISCA.Presentation.Controls.MsgBox.Show("資料驗證失敗，請修正後再行儲存", "驗證失敗", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            DSXmlHelper InsertHelper = new DSXmlHelper("InsertRequest");
            DSXmlHelper updateHelper = new DSXmlHelper("UpdateRequest");
            List<string> deleteList = new List<string>();
            //ISemester semester = SemesterProvider.GetInstance();
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                StudentRowTag tag = row.Tag as StudentRowTag;
                //semester.SetDate(tag.Date);

                //log 紀錄修改後的資料 日期部分
                if (!afterData.ContainsKey(tag.Student.ID))
                    afterData.Add(tag.Student.ID, new Dictionary<string, string>());

                if (tag.RowID == null)
                {
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

                        //log 紀錄修改後的資料 缺曠明細部分
                        if (!afterData[tag.Student.ID].ContainsKey(pinfo.Name))
                            afterData[tag.Student.ID].Add(pinfo.Name, ainfo.Name);
                    }
                    if (hasContent)
                    {
                        InsertHelper.AddElement("Attendance");
                        InsertHelper.AddElement("Attendance", "Field");
                        InsertHelper.AddElement("Attendance/Field", "RefStudentID", tag.Student.ID);
                        InsertHelper.AddElement("Attendance/Field", "SchoolYear", row.Cells[ColumnIndex["學年度"]].Value.ToString());
                        InsertHelper.AddElement("Attendance/Field", "Semester", row.Cells[ColumnIndex["學期"]].Value.ToString());
                        InsertHelper.AddElement("Attendance/Field", "OccurDate", dateTimeInput1.Value.ToShortDateString());
                        InsertHelper.AddElement("Attendance/Field", "Detail", h2.GetRawXml(), true);
                    }

                }
                else // 若是原本就有紀錄的
                {
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

                        //log 紀錄修改後的資料 缺曠明細部分
                        if (!afterData[tag.Student.ID].ContainsKey(pinfo.Name))
                            afterData[tag.Student.ID].Add(pinfo.Name, ainfo.Name);
                    }

                    if (hasContent)
                    {
                        updateHelper.AddElement("Attendance");
                        updateHelper.AddElement("Attendance", "Field");
                        updateHelper.AddElement("Attendance/Field", "RefStudentID", tag.Student.ID);
                        updateHelper.AddElement("Attendance/Field", "SchoolYear", row.Cells[ColumnIndex["學年度"]].Value.ToString());
                        updateHelper.AddElement("Attendance/Field", "Semester", row.Cells[ColumnIndex["學期"]].Value.ToString());
                        updateHelper.AddElement("Attendance/Field", "OccurDate", dateTimeInput1.Value.ToShortDateString());
                        updateHelper.AddElement("Attendance/Field", "Detail", h2.GetRawXml(), true);
                        updateHelper.AddElement("Attendance", "Condition");
                        updateHelper.AddElement("Attendance/Condition", "ID", tag.RowID);
                    }
                    else
                    {
                        deleteList.Add(tag.RowID);

                        //log 紀錄被刪除的資料
                        afterData.Remove(tag.Student.ID);
                        deleteData.Add(tag.Student.ID);
                    }
                }
            }
            if (InsertHelper.GetElements("Attendance").Length > 0)
            {
                #region 新增
                try
                {
                    EditAttendance.Insert(new DSRequest(InsertHelper));
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("缺曠紀錄新增失敗 : " + ex.Message, "新增失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //log 寫入log
                foreach (string studentid in afterData.Keys)
                {
                    if (!beforeData.ContainsKey(studentid) && afterData[studentid].Count > 0)
                    {
                        K12.Data.StudentRecord sr = K12.Data.Student.SelectByID(studentid);

                        StringBuilder desc = new StringBuilder("");
                        desc.AppendLine("學生「" + sr.Name + "」");
                        desc.AppendLine("日期「" + logDate.ToShortDateString() + "」");
                        foreach (string period in afterData[studentid].Keys)
                        {
                            desc.AppendLine("節次「" + period + "」為「" + afterData[studentid][period] + "」 ");
                        }
                        //Log部份
                        //CurrentUser.Instance.AppLog.Write(EntityType.Student, EntityAction.Insert, studentid, desc.ToString(), this.Text, "");

                        ApplicationLog.Log("學務系統.缺曠資料", "批次新增缺曠資料", "student", sr.ID, desc.ToString());
                    }
                }
                #endregion
            }
            if (updateHelper.GetElements("Attendance").Length > 0)
            {
                #region 修改
                try
                {
                    EditAttendance.Update(new DSRequest(updateHelper));
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("缺曠紀錄更新失敗 : " + ex.Message, "更新失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //log 寫入log
                foreach (string studentid in afterData.Keys)
                {
                    if (beforeData.ContainsKey(studentid) && afterData[studentid].Count > 0)
                    {
                        K12.Data.StudentRecord sr = K12.Data.Student.SelectByID(studentid);
                        bool dirty = false;
                        StringBuilder desc = new StringBuilder("");
                        desc.AppendLine("學生「" + sr.Name + "」");
                        desc.AppendLine("日期「" + logDate.ToShortDateString() + "」");
                        foreach (string period in beforeData[studentid].Keys)
                        {
                            if (!afterData[studentid].ContainsKey(period))
                                afterData[studentid].Add(period, "");
                        }
                        foreach (string period in afterData[studentid].Keys)
                        {
                            if (beforeData[studentid].ContainsKey(period))
                            {
                                if (beforeData[studentid][period] != afterData[studentid][period])
                                {
                                    dirty = true;
                                    desc.AppendLine("節次「" + period + "」由「" + beforeData[studentid][period] + "」變更為「" + afterData[studentid][period] + "」 ");
                                }
                            }
                            else
                            {
                                dirty = true;
                                desc.AppendLine("節次「" + period + "」為「" + afterData[studentid][period] + "」 ");
                            }

                        }
                        if (dirty)
                        {
                            //Log部份
                            //CurrentUser.Instance.AppLog.Write(EntityType.Student, EntityAction.Update, studentid, desc.ToString(), this.Text, "");
                            ApplicationLog.Log("學務系統.缺曠資料", "批次修改缺曠資料", "student", sr.ID, desc.ToString());
                        }

                    }
                }
                #endregion
            }
            if (deleteList.Count > 0)
            {
                #region 刪除
                DSXmlHelper deleteHelper = new DSXmlHelper("DeleteRequest");
                deleteHelper.AddElement("Attendance");
                foreach (string key in deleteList)
                {
                    deleteHelper.AddElement("Attendance", "ID", key);
                }

                try
                {
                    EditAttendance.Delete(new DSRequest(deleteHelper));
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("缺曠紀錄刪除失敗 : " + ex.Message, "刪除失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //log 寫入被刪除的資料的log
                foreach (string studentid in deleteData)
                {
                    K12.Data.StudentRecord sr = K12.Data.Student.SelectByID(studentid);
                    StringBuilder desc = new StringBuilder("");
                    desc.AppendLine("學生「" + sr.Name + "」");
                    desc.AppendLine("刪除「" + logDate.ToShortDateString() + "」缺曠紀錄 ");
                    //Log部份
                    //CurrentUser.Instance.AppLog.Write(EntityType.Student, EntityAction.Delete, studentid, desc.ToString(), this.Text, "");
                    ApplicationLog.Log("學務系統.缺曠資料", "批次刪除缺曠資料", "student", sr.ID, desc.ToString());
                }
                #endregion
            }

            List<string> studentids = new List<string>();
            foreach (K12.Data.StudentRecord var in _students)
            {
                studentids.Add(var.ID);
            }
            if (studentids.Count > 0)
            {
                Attendance.Instance.SyncDataBackground(studentids.ToArray());

                //JHAttendance.SyncData();
            }
            FISCA.Presentation.Controls.MsgBox.Show("儲存缺曠資料成功!", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();

            SaveDateSetting();       //儲存日期是否鎖定的設定 
            #endregion
        }

        private bool IsValid()
        {
            #region DataGridView資料驗證(如果ErrorText內容為空)

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.ErrorText))
                        return false;
                }
            }
            return true; 
            #endregion
        }

        private void btnRenew_Click(object sender, EventArgs e)
        {
            //if (!startDate.IsValid)
            //{
            //    _errProvider.SetError(startDate, "日期格式錯誤");
            //    return;
            //}
            SearchStudentRange();
            GetAbsense();
            chkHasData_CheckedChanged(null, null);
        }

        private void dataGridView_KeyDown(object sender, KeyEventArgs e)
        {
            #region 如果按下按鈕
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
            #region 學年度/學期輸入驗證
            DataGridViewCell cell = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
            DataGridViewColumn column = dataGridView.Columns[e.ColumnIndex];
            if (column.Index == ColumnIndex["學年度"])
            {
                string errorMessage = "";
                int schoolYear;
                if (cell.Value == null)
                    errorMessage = "學年度不可為空白";
                else if (!int.TryParse(cell.Value.ToString(), out schoolYear))
                    errorMessage = "學年度必須為整數";

                if (errorMessage != "")
                {
                    cell.ErrorText = errorMessage;
                    //cell.Style.BackColor = Color.Red;
                    //cell.ToolTipText = errorMessage;
                }
                else
                {
                    cell.ErrorText = string.Empty;
                    SemesterCellInfo cinfo = cell.Tag as SemesterCellInfo;
                    cinfo.SetValue(cell.Value == null ? string.Empty : cell.Value.ToString());
                    //cell.Style.BackColor = Color.White;
                    //cell.ToolTipText = "";
                }
            }
            else if (column.Index == ColumnIndex["學期"])
            {
                string errorMessage = string.Empty;

                if (cell.Value == null)
                    errorMessage = "學期不可為空白";
                else if (cell.Value.ToString() != "1" && cell.Value.ToString() != "2")
                    errorMessage = "學期必須為整數『1』或『2』";

                if (errorMessage != string.Empty)
                {
                    cell.ErrorText = errorMessage;
                    //cell.Style.BackColor = Color.Red;
                    //cell.ToolTipText = errorMessage;
                }
                else
                {
                    cell.ErrorText = string.Empty;
                    SemesterCellInfo cinfo = cell.Tag as SemesterCellInfo;
                    cinfo.SetValue(cell.Value == null ? string.Empty : cell.Value.ToString());
                    //cell.Style.BackColor = Color.White;
                    //cell.ToolTipText = "";
                }
            } 
            #endregion
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            if (IsDirty())
            {
                if (FISCA.Presentation.Controls.MsgBox.Show("資料已變更且尚未儲存，是否放棄已編輯資料?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    this.Close();
                }
            }
            else
                this.Close();
        }

        private void dateTimeInput1_Validated(object sender, EventArgs e)
        {
            if (IsDirty())
            {
                if (FISCA.Presentation.Controls.MsgBox.Show("資料已變更且尚未儲存，是否放棄已編輯資料?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    dateTimeInput1.Value = _currentDate;
                    return;
                }
            }
            _currentDate = dateTimeInput1.Value;
            SearchStudentRange();
            GetAbsense();
            chkHasData_CheckedChanged(null, null);
        }
        
        private void picLock_Click(object sender, EventArgs e)
        {
            //bool isLock = false;
            //if (picLock.Tag != null)
            //{
            //    if (!bool.TryParse(picLock.Tag.ToString(), out isLock))
            //        isLock = false;
            //}
            //if (isLock)
            //{
            //    picLock.Image = Resources.unlock;
            //    picLock.Tag = false;
            //    toolTip.SetToolTip(picLock, "缺曠日期為未鎖定狀態，您可以點選圖示，將特定日期鎖定。");
            //    labelX2.Text = "";
            //}
            //else
            //{
            //    picLock.Image = Resources._lock;
            //    picLock.Tag = true;
            //    toolTip.SetToolTip(picLock, "缺曠日期已鎖定，您可以點選圖示解除鎖定。");
            //    labelX2.Text = "已鎖定缺曠日期";
            //}

            //this.SaveDateSetting();
        }

        private bool IsDirty()
        { 
            #region 資料驗證
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Tag == null) continue; //Tag是空的
                    if (cell.Tag is SemesterCellInfo) //學年度學期
                    {
                        SemesterCellInfo cInfo = cell.Tag as SemesterCellInfo;
                        if (cInfo.IsDirty) return true;
                    }
                    else if (cell.Tag is AbsenceCellInfo) //缺曠別
                    {
                        AbsenceCellInfo cInfo = cell.Tag as AbsenceCellInfo;
                        if (cInfo.IsDirty) return true;
                    }
                }
            }
            return false; 
            #endregion
        }

        private void chkHasData_CheckedChanged(object sender, EventArgs e)
        {
            #region 僅顯示有缺曠的資料
            dataGridView.SuspendLayout();

            if (chkHasData.Checked == true)
            {
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
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
                    if (hasData == false)
                    {
                        _hiddenRows.Add(row);
                        row.Visible = false;
                    }
                }
            }
            else
            {
                foreach (DataGridViewRow row in _hiddenRows)
                    row.Visible = true;
            }

            dataGridView.ResumeLayout(); 
            #endregion
        }

        private void dateTimeInput1_TextChanged(object sender, EventArgs e)
        {
            if (dataGridView.Rows.Count != 0)
            {
                btnRenew_Click(null, null);
            }
        }

        //關閉視窗即儲存設定
        private void MutiEditor_FormClosing(object sender, FormClosingEventArgs e)
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

    public class StudentRowTag
    {
        private K12.Data.StudentRecord _student;

        public K12.Data.StudentRecord Student
        {
            get { return _student; }
            set { _student = value; }
        }
        private string _RowID;

        public string RowID
        {
            get { return _RowID; }
            set { _RowID = value; }
        }
    }
}