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
        private Dictionary<string, AbsenceInfo> _absenceList;
        private StudentRecord _student;
        private ISemester _semesterProvider;
        private int _startIndex;
        private ErrorProvider _errorProvider;
        private DateTime _currentStartDate;
        private DateTime _currentEndDate;

        Dictionary<string, int> ColumnIndex = new Dictionary<string, int>();

        List<string> WeekDay = new List<string>();
        List<DayOfWeek> nowWeekDay = new List<DayOfWeek>();
        List<DateTime> _Holidays = new List<DateTime>();

        //System.Windows.Forms.ToolTip toolTip = new System.Windows.Forms.ToolTip();

        //log 需要用到的
        private Dictionary<string, Dictionary<string, string>> beforeData = new Dictionary<string, Dictionary<string, string>>();
        private Dictionary<string, Dictionary<string, string>> afterData = new Dictionary<string, Dictionary<string, string>>();
        private List<string> deleteData = new List<string>();

        public SingleEditor(StudentRecord student)
        {
            InitializeComponent(); //設計工具產生的

            _errorProvider = new ErrorProvider();
            _student = student;
            _absenceList = new Dictionary<string, AbsenceInfo>();
            _semesterProvider = SemesterProvider.GetInstance();
            //_startIndex = 4;

        }

        private void SingleEditor_Load(object sender, EventArgs e)
        {
            #region Load
            this.Text = "單人多天缺曠管理";
            StringBuilder sb = new StringBuilder();
            string ClassName = _student.Class != null ? _student.Class.Name : "";
            sb.Append("班級：<b>" + ClassName + "</b>　");
            sb.Append("座號：<b>" + _student.SeatNo + "</b>　");
            sb.Append("姓名：<b>" + _student.Name + "</b>　");
            sb.Append("學號：<b>" + _student.StudentNumber + "</b>");
            lblInfo.Text = sb.ToString();
            InitializeRadioButton();
            InitializeDateRange(); //取得日期定義
            InitializeDataGridViewColumn();
            //SearchDateRange();
            //GetAbsense();
            LoadAbsense();

            #endregion
        }

        private void InitializeDateRange()
        {
            #region 日期定義
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
                DateConfig.Save(); //儲存此預設檔
            }

            XmlElement loadXml = DSXmlHelper.LoadXml(date);
            checkBoxX1.Checked = bool.Parse(loadXml.SelectSingleNode("Locked").InnerText);

            if (checkBoxX1.Checked) //如果是鎖定,就取鎖定日期
            {
                dateTimeInput1.Text = loadXml.SelectSingleNode("StartDate").InnerText;
                dateTimeInput2.Text = loadXml.SelectSingleNode("EndDate").InnerText;
            }
            else //如果沒有鎖定,就取當天
            {
                dateTimeInput1.Text = DateTime.Today.AddDays(-6).ToShortDateString();
                dateTimeInput2.Text = DateTime.Today.ToShortDateString();
            }
            _currentStartDate = dateTimeInput1.Value;
            _currentEndDate = dateTimeInput2.Value;
            #endregion
        }

        private void SaveDateSetting()
        {
            #region 儲存日期資料
            K12.Data.Configuration.ConfigData DateConfig = K12.Data.School.Configuration["Attendance_BatchEditor"];

            DSXmlHelper helper = new DSXmlHelper("xml");
            helper.AddElement("StartDate");
            helper.AddText("StartDate", dateTimeInput1.Value.ToShortDateString());
            helper.AddElement("EndDate");
            helper.AddText("EndDate", dateTimeInput2.Value.ToShortDateString());
            helper.AddElement("Locked");
            helper.AddText("Locked", checkBoxX1.Checked.ToString());

            DateConfig["SingleEditor"] = helper.BaseElement.OuterXml;
            DateConfig.Save(); //儲存此預設檔

            #endregion
        }

        private void InitializeRadioButton()
        {
            #region 缺曠類別建立
            DSResponse dsrsp = Config.GetAbsenceList();
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

        private void InitializeDataGridViewColumn()
        {
            #region DataGridView的Column建立

            ColumnIndex.Clear(); //清除

            DSResponse dsrsp = Config.GetPeriodList();
            DSXmlHelper helper = dsrsp.GetContent();
            PeriodCollection collection = new PeriodCollection();
            foreach (XmlElement element in helper.GetElements("Period"))
            {
                PeriodInfo info = new PeriodInfo(element);
                collection.Items.Add(info);
            }
            int ColumnsIndex = dataGridView.Columns.Add("colDate", "日期");
            ColumnIndex.Add("日期", ColumnsIndex);
            dataGridView.Columns[ColumnsIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView.Columns[ColumnsIndex].ReadOnly = true;

            ColumnsIndex = dataGridView.Columns.Add("colWeek", "星期");
            ColumnIndex.Add("星期", ColumnsIndex);
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
            dataGridView.Columns[ColumnsIndex].Frozen = true;

            _startIndex = ColumnIndex["學期"] + 1;

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
            #region 日期選擇
            DateTime start = dateTimeInput1.Value;
            DateTime end = dateTimeInput2.Value;

            DateTime date = start;
            dataGridView.Rows.Clear();

            TimeSpan ts = dateTimeInput2.Value - dateTimeInput1.Value;
            if (ts.Days > 1500)
            {
                FISCA.Presentation.Controls.MsgBox.Show("您選取了" + ts.Days.ToString() + "天\n由於選取日期區間過長,請重新設定日期！");
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

                row.Cells[ColumnIndex["日期"]].Value = dateValue;
                row.Cells[ColumnIndex["星期"]].Value = GetDayOfWeekInChinese(date.DayOfWeek);
                _semesterProvider.SetDate(date);
                row.Cells[ColumnIndex["學年度"]].Value = _semesterProvider.SchoolYear;
                row.Cells[ColumnIndex["學期"]].Value = _semesterProvider.Semester;
                date = date.AddDays(1);

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
            helper.AddElement("Condition", "RefStudentID", _student.ID);
            helper.AddElement("Condition", "StartDate", dateTimeInput1.Value.ToShortDateString());
            helper.AddElement("Condition", "EndDate", dateTimeInput2.Value.ToShortDateString());
            DSResponse dsrsp = JHSchool.Feature.Legacy.QueryAttendance.GetAttendance(new DSRequest(helper));
            helper = dsrsp.GetContent();

            //log 清除 beforeData
            beforeData.Clear();

            foreach (XmlElement element in helper.GetElements("Attendance"))
            {
                // 這裡要做一些事情  例如找到東家塞進去
                string occurDate = element.SelectSingleNode("OccurDate").InnerText;
                string schoolYear = element.SelectSingleNode("SchoolYear").InnerText;
                string semester = element.SelectSingleNode("Semester").InnerText;
                string id = element.GetAttribute("ID");
                XmlNode dNode = element.SelectSingleNode("Detail").FirstChild;

                //log 紀錄修改前的資料 日期部分
                DateTime logDate;
                if (DateTime.TryParse(occurDate, out logDate))
                {
                    if (!beforeData.ContainsKey(logDate.ToShortDateString()))
                        beforeData.Add(logDate.ToShortDateString(), new Dictionary<string, string>());
                }

                DataGridViewRow row = null;
                foreach (DataGridViewRow r in dataGridView.Rows)
                {
                    DateTime date;
                    RowTag rt = r.Tag as RowTag;

                    if (!DateTime.TryParse(occurDate, out date)) continue;
                    if (date.CompareTo(rt.Date) != 0) continue;
                    row = r;
                }

                if (row == null) continue;
                RowTag rowTag = row.Tag as RowTag;
                rowTag.IsNew = false;
                rowTag.Key = id;

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
                            if (!beforeData[logDate.ToShortDateString()].ContainsKey(info.Name))
                                beforeData[logDate.ToShortDateString()].Add(info.Name, ai.Name);

                            break;
                        }
                    }
                }
            }
            #endregion
        }

        //儲存
        private void btnSave_Click(object sender, EventArgs e)
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

            List<string> synclist = new List<string>();

            ISemester semester = SemesterProvider.GetInstance();
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                RowTag tag = row.Tag as RowTag;
                semester.SetDate(tag.Date);

                //log 紀錄修改後的資料 日期部分
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

                        //log 紀錄修改後的資料 缺曠明細部分
                        if (!afterData[tag.Date.ToShortDateString()].ContainsKey(pinfo.Name))
                            afterData[tag.Date.ToShortDateString()].Add(pinfo.Name, ainfo.Name);

                    }
                    if (hasContent)
                    {
                        InsertHelper.AddElement("Attendance");
                        InsertHelper.AddElement("Attendance", "Field");
                        InsertHelper.AddElement("Attendance/Field", "RefStudentID", _student.ID);
                        InsertHelper.AddElement("Attendance/Field", "SchoolYear", row.Cells[ColumnIndex["學年度"]].Value.ToString());
                        InsertHelper.AddElement("Attendance/Field", "Semester", row.Cells[ColumnIndex["學期"]].Value.ToString());
                        InsertHelper.AddElement("Attendance/Field", "OccurDate", tag.Date.ToShortDateString());
                        InsertHelper.AddElement("Attendance/Field", "Detail", h2.GetRawXml(), true);
                    }

                    #endregion
                }
                else // 若是原本就有紀錄的
                {
                    #region 是舊的
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
                        if (!afterData[tag.Date.ToShortDateString()].ContainsKey(pinfo.Name))
                            afterData[tag.Date.ToShortDateString()].Add(pinfo.Name, ainfo.Name);
                    }

                    if (hasContent)
                    {
                        updateHelper.AddElement("Attendance");
                        updateHelper.AddElement("Attendance", "Field");
                        updateHelper.AddElement("Attendance/Field", "RefStudentID", _student.ID);
                        updateHelper.AddElement("Attendance/Field", "SchoolYear", row.Cells[ColumnIndex["學年度"]].Value.ToString());
                        updateHelper.AddElement("Attendance/Field", "Semester", row.Cells[ColumnIndex["學期"]].Value.ToString());
                        updateHelper.AddElement("Attendance/Field", "OccurDate", tag.Date.ToShortDateString());
                        updateHelper.AddElement("Attendance/Field", "Detail", h2.GetRawXml(), true);
                        updateHelper.AddElement("Attendance", "Condition");
                        updateHelper.AddElement("Attendance/Condition", "ID", tag.Key);
                    }
                    else
                    {
                        deleteList.Add(tag.Key);

                        //log 紀錄被刪除的資料
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
                    FISCA.Presentation.Controls.MsgBox.Show("缺曠紀錄新增失敗 : " + ex.Message, "新增失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


                //log 寫入log
                foreach (string date in afterData.Keys)
                {
                    if (!beforeData.ContainsKey(date) && afterData[date].Count > 0)
                    {
                        StringBuilder desc = new StringBuilder("");
                        desc.AppendLine("學生「" + Student.Instance.Items[_student.ID].Name + "」");
                        desc.AppendLine("日期「" + date + "」");
                        foreach (string period in afterData[date].Keys)
                        {
                            desc.AppendLine("節次「" + period + "」設為「" + afterData[date][period] + "」");
                        }

                        ApplicationLog.Log("學務系統.缺曠資料", "批次新增缺曠資料", "student", _student.ID, desc.ToString());
                        //Log部份
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
                    FISCA.Presentation.Controls.MsgBox.Show("缺曠紀錄更新失敗 : " + ex.Message, "更新失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //log 寫入log
                foreach (string date in afterData.Keys)
                {
                    if (beforeData.ContainsKey(date) && afterData[date].Count > 0)
                    {
                        bool dirty = false;
                        StringBuilder desc = new StringBuilder("");
                        desc.AppendLine("學生「" + Student.Instance.Items[_student.ID].Name + "」 ");
                        desc.AppendLine("日期「" + date + "」 ");
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
                                    desc.AppendLine("節次「" + period + "」由「" + beforeData[date][period] + "」變更為「" + afterData[date][period] + "」");
                                }
                            }
                            else
                            {
                                dirty = true;
                                desc.AppendLine("節次「" + period + "」由「」變更為「" + afterData[date][period] + "」 ");
                            }

                        }
                        if (dirty)
                        {
                            //Log部份
                            ApplicationLog.Log("學務系統.缺曠資料", "批次修改缺曠資料", "student", _student.ID, desc.ToString());
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
                    FISCA.Presentation.Controls.MsgBox.Show("缺曠紀錄刪除失敗 : " + ex.Message, "刪除失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //log 寫入被刪除的資料的log
                StringBuilder desc = new StringBuilder("");
                desc.AppendLine("學生「" + Student.Instance.Items[_student.ID].Name + "」");
                foreach (string date in deleteData)
                {
                    desc.AppendLine("刪除「" + date + "」缺曠紀錄 ");
                }
                //Log部份
                ApplicationLog.Log("學務系統.缺曠資料", "批次刪除缺曠資料", "student", _student.ID, desc.ToString());
                //CurrentUser.Instance.AppLog.Write(EntityType.Student, EntityAction.Delete, _student.ID, desc.ToString(), this.Text, "");
            }
            #endregion

            //觸發變更事件
            Attendance.Instance.SyncDataBackground(_student.ID);
            //Student.Instance.SyncDataBackground(_student.ID);

            FISCA.Presentation.Controls.MsgBox.Show("儲存缺曠資料成功!", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
            #endregion
        }

        private bool IsValid()
        {
            #region DataGridView資料驗證(如果ErrorText內容為空)
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
            LoadAbsense(); //重新整理
        }

        private void LoadAbsense()
        {
            #region 處理星期設定檔
            WeekDay.Clear();
            nowWeekDay.Clear();
            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["缺曠批次登錄_星期設定"];
            string cdIN = cd["星期設定"];

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
                WeekDay.AddRange(new string[] { "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日" });
            }

            nowWeekDay = ChengDayOfWeel(WeekDay);
            #endregion

            #region 取得假別清單
            ConfigData _CD;
            //取得之前設定設定
            _CD = School.Configuration["SCHOOL_HOLIDAY_CONFIG_STRING"];
            XElement rootXml = null;
            string xmlContent = _CD["CONFIG_STRING"];

            if (!string.IsNullOrWhiteSpace(xmlContent))
                rootXml = XElement.Parse(xmlContent);
            else
                rootXml = new XElement("SchoolHolidays");

            //讀出之前的假日清單
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
            #region 取得星期對照表
            List<DayOfWeek> DOW = new List<DayOfWeek>();
            foreach (string each in list)
            {
                if (each == "星期一")
                {
                    DOW.Add(DayOfWeek.Monday);
                }
                else if (each == "星期二")
                {
                    DOW.Add(DayOfWeek.Tuesday);
                }
                else if (each == "星期三")
                {
                    DOW.Add(DayOfWeek.Wednesday);
                }
                else if (each == "星期四")
                {
                    DOW.Add(DayOfWeek.Thursday);
                }
                else if (each == "星期五")
                {
                    DOW.Add(DayOfWeek.Friday);
                }
                else if (each == "星期六")
                {
                    DOW.Add(DayOfWeek.Saturday);
                }
                else if (each == "星期日")
                {
                    DOW.Add(DayOfWeek.Sunday);
                }
            }

            return DOW;
            #endregion
        }

        private string GetDayOfWeekInChinese(DayOfWeek day)
        {
            #region 星期(中/英)對照表
            switch (day)
            {
                case DayOfWeek.Monday:
                    return "一";
                case DayOfWeek.Tuesday:
                    return "二";
                case DayOfWeek.Wednesday:
                    return "三";
                case DayOfWeek.Thursday:
                    return "四";
                case DayOfWeek.Friday:
                    return "五";
                case DayOfWeek.Saturday:
                    return "六";
                default:
                    return "日";
            }
            #endregion
        }

        private void btnClose_Click(object sender, EventArgs e)
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
            if (e.ColumnIndex == ColumnIndex["學年度"])
            {
                string errorMessage = "";
                int schoolYear;
                if (cell.Value == null)
                    errorMessage = "學年度不可為空白";
                else if (!int.TryParse(cell.Value.ToString(), out schoolYear))
                    errorMessage = "學年度必須為整數";

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
            else if (e.ColumnIndex == ColumnIndex["學期"])
            {
                string errorMessage = "";

                if (cell.Value == null)
                    errorMessage = "學期不可為空白";
                else if (cell.Value.ToString() != "1" && cell.Value.ToString() != "2")
                    errorMessage = "學期必須為整數『1』或『2』";

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
        //        toolTip.SetToolTip(picLock, "缺曠登錄日期為未鎖定狀態，您可以點選圖示，將特定日期區間鎖定。");
        //        labelX2.Text = "";
        //    }
        //    else
        //    {
        //        picLock.Image = Resources._lock;
        //        picLock.Tag = true;
        //        toolTip.SetToolTip(picLock, "缺曠登錄日期已鎖定，您可以點選圖示解除鎖定。");
        //        labelX2.Text = "已鎖定缺曠日期";
        //    }
        //    SaveDateSetting();
        //}

        private void dateTimeInput1_Validated(object sender, EventArgs e)
        {
            #region dateTimeInput1資料變更事件
            _errorProvider.SetError(dateTimeInput1, string.Empty);

            if (IsDirty())
            {
                if (FISCA.Presentation.Controls.MsgBox.Show("資料已變更且尚未儲存，是否放棄已編輯資料?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
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
            #region dateTimeInput1資料變更事件
            _errorProvider.SetError(dateTimeInput2, string.Empty);

            if (IsDirty())
            {
                if (FISCA.Presentation.Controls.MsgBox.Show("資料已變更且尚未儲存，是否放棄已編輯資料?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
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
            #region 資料驗證
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
                if (!nowWeekDay.Contains(rowTag.Date.DayOfWeek)) //篩選不顯示的星期
                {
                    isHolday = true;
                }
                else if (_Holidays.Contains(rowTag.Date)) //篩選假日
                {
                    isHolday = true;
                }

                if (hasData)
                {
                    if (isHolday)
                    {
                        row.Cells[ColumnIndex["星期"]].Style.ForeColor = Color.Red;
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
            #region 星期設定
            Searchday Sday = new Searchday("缺曠批次登錄_星期設定");
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