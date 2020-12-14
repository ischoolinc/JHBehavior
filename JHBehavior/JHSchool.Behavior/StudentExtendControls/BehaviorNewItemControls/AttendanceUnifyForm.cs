using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using JHSchool.Data;
using JHSchool.Behavior.BusinessLogic;
using JHSchool.Behavior.Feature;
using JHSchool.Editor;
using Framework.Security;
using Framework;
using FISCA.LogAgent;
using FISCA.DSAUtil;
using System.Xml;
using Campus.Windows;

namespace JHSchool.Behavior.StudentExtendControls
{
    public partial class AttendanceUnifyForm : BaseForm
    {
        private FeatureAce _UserPermission;

        private BackgroundWorker BGW = new BackgroundWorker();
        private bool BkWBool = false;

        private  string _StudentID;
        private int _SchoolYear;
        private int _Semester;

        private List<JHAttendanceRecord> AttendanceList = new List<JHAttendanceRecord>();
        private List<AutoSummaryRecord> AutoRecord = new List<AutoSummaryRecord>();
        private JHMoralScoreRecord MSRecord;

        private List<K12.Data.PeriodMappingInfo> _periodList = new List<K12.Data.PeriodMappingInfo>();
        private List<K12.Data.AbsenceMappingInfo> _absenceList = new List<K12.Data.AbsenceMappingInfo>();

        private List<string> _periodTypes = new List<string>();
        private List<string> _absenceTypes = new List<string>();

        private const string SCHOOL_YEAR_COLUMN = "學年度";
        private const string SEMESTER_COLUMN = "學期";

        //畫面顏色
        private Color SetColor = Color.LightCyan;

        //缺曠定位
        Dictionary<string, int> AttendanceIndex = new Dictionary<string, int>();

        //明細缺曠加總
        Dictionary<string, int> AttendanceCount = new Dictionary<string, int>();

        Dictionary<string, string> PeriodDic = new Dictionary<string, string>(); //節次對照表

        public AttendanceUnifyForm(string StudentID, int SchoolYear, int Semester, FeatureAce UserPermission)
        {
            InitializeComponent();

            _StudentID = StudentID;
            _SchoolYear = SchoolYear;
            _Semester = Semester;
            _UserPermission = UserPermission;
        }

        private void AttendanceUnifyForm_Load(object sender, EventArgs e)
        {

            lbHelp1.Text = _SchoolYear.ToString() + "學年度　第" + _Semester.ToString() + "學期　缺曠記錄";

            BGW.DoWork += new DoWorkEventHandler(BgW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BgW_RunWorkerCompleted);

            JHAttendance.AfterDelete += new EventHandler<K12.Data.DataChangedEventArgs>(DataChanged);
            JHAttendance.AfterInsert += new EventHandler<K12.Data.DataChangedEventArgs>(DataChanged);
            JHAttendance.AfterUpdate += new EventHandler<K12.Data.DataChangedEventArgs>(DataChanged);

            JHMoralScore.AfterDelete += new EventHandler<K12.Data.DataChangedEventArgs>(DataChanged);
            JHMoralScore.AfterInsert += new EventHandler<K12.Data.DataChangedEventArgs>(DataChanged);
            JHMoralScore.AfterUpdate += new EventHandler<K12.Data.DataChangedEventArgs>(DataChanged);

            BGW.RunWorkerAsync();

        }

        void DataChanged(object sender, K12.Data.DataChangedEventArgs e)
        {
            if (BGW.IsBusy)
            {
                BkWBool = true;
            }
            else
            {
                lockForm(false);
                BGW.RunWorkerAsync();
            }
        }

        private void lockForm(bool IsTrueOrFalse)
        {

            btnAttendanceNew.Enabled = IsTrueOrFalse;
            btnAttendanceEdit.Enabled = IsTrueOrFalse;
            btnAttendanceDelete.Enabled = IsTrueOrFalse;
            btnSaveAttendanceStatistics.Enabled = IsTrueOrFalse;
            dgvAttendance.Enabled = IsTrueOrFalse;
            if (IsTrueOrFalse)
            {
                this.Text = "缺曠學期統計";
            }
            else
            {
                this.Text = "更新資料中...";
            }
        }

        void BgW_DoWork(object sender, DoWorkEventArgs e)
        {
            AttendanceList = JHAttendance.SelectBySchoolYearAndSemester(JHStudent.SelectByIDs(new string[] { _StudentID }), _SchoolYear, _Semester);

            //取得自動統計
            SchoolYearSemester SchoolSemester = new SchoolYearSemester(_SchoolYear, _Semester);

            AutoRecord = AutoSummary.Select(new string[] { _StudentID }, new SchoolYearSemester[] { SchoolSemester });

            //取得日常生活表現
            MSRecord = JHMoralScore.SelectBySchoolYearAndSemester(_StudentID, _SchoolYear, _Semester);
        }

        void BgW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (BkWBool)
            {
                BkWBool = false;
                BGW.RunWorkerAsync();
                return;
            }

            lockForm(true);

            //建立畫面內容
            Initialize();

            BingDataAttendance();

            BingDataMoralScore();
        }

        //建立預設畫面
        private void Initialize()
        {
            #region 缺曠明細
            listViewAttendance.Items.Clear();

            _periodList.Clear();
            _periodList = K12.Data.PeriodMapping.SelectAll();
            _absenceList.Clear();
            _absenceList = K12.Data.AbsenceMapping.SelectAll();

            listViewAttendance.Columns.Clear();
            listViewAttendance.Columns.Add("OccurDate", "缺曠日期");
            listViewAttendance.Columns.Add("DayOfWeek", "星期");

            PeriodDic.Clear(); //節次對照表

            foreach (K12.Data.PeriodMappingInfo info in _periodList)
            {
                //節次對照表
                if (!PeriodDic.ContainsKey(info.Name))
                {
                    PeriodDic.Add(info.Name, info.Type);
                }

                ColumnHeader column = listViewAttendance.Columns.Add(info.Name, info.Name);
                column.Tag = info;
            } 
            #endregion

            #region 缺曠統計
            _periodTypes.Clear();
            _periodTypes = GetPeriodTypeItems();
            _periodTypes.Sort();
            _absenceTypes.Clear();
            _absenceTypes = GetAbsenceItems();

            dgvAttendance.Columns.Clear();

            AttendanceIndex.Clear();
            AttendanceCount.Clear();
            int sortEnd = dgvAttendance.Columns.Add("統計類型", "統計類型");            
            dgvAttendance.Columns[sortEnd].SortMode = DataGridViewColumnSortMode.NotSortable;

            List<string> cols1 = new List<string>();

            foreach (string periodType in _periodTypes)
            {
                foreach (string each in _absenceTypes)
                {
                    string columnName = periodType + each;

                    cols1.Add(columnName);
                    AttendanceIndex.Add(columnName, dgvAttendance.Columns.Add(columnName, columnName));
                    dgvAttendance.Columns[columnName].Width = 80;
                    dgvAttendance.Columns[columnName].Tag = periodType + ":" + each;
                    dgvAttendance.Columns[columnName].SortMode = DataGridViewColumnSortMode.NotSortable;
                    AttendanceCount.Add(columnName, 0);
                }
            }
            DataGridViewImeDecorator dec1 = new DataGridViewImeDecorator(this.dgvAttendance, cols1);
            #endregion
        }

        //缺曠資料
        private void BingDataAttendance()
        {
            //排序
            AttendanceList.Sort(new Comparison<JHAttendanceRecord>(SchoolYearComparer));

            foreach (JHAttendanceRecord record in AttendanceList)
            {
                ListViewItem lvItem = listViewAttendance.Items.Add(record.OccurDate.ToShortDateString());
                lvItem.SubItems.Add(record.DayOfWeek);

                lvItem.Tag = record;

                if (listViewAttendance.Columns.Count != 0)
                {
                    listViewAttendance.Columns[0].Width = 110;
                }

                for (int i = 2; i < listViewAttendance.Columns.Count; i++)
                    lvItem.SubItems.Add("");

                for (int i = 2; i < listViewAttendance.Columns.Count; i++)
                {
                    ColumnHeader column = listViewAttendance.Columns[i];
                    K12.Data.PeriodMappingInfo info = column.Tag as K12.Data.PeriodMappingInfo;

                    //if (record.PeriodDetail == null) continue;

                    foreach (K12.Data.AttendancePeriod period in record.PeriodDetail)
                    {
                        if (info == null) continue;
                        if (period.Period != info.Name) continue;
                        if (period.AbsenceType == null) continue;

                        if (PeriodDic.ContainsKey(period.Period))
                        {
                            string periodName = PeriodDic[period.Period] + period.AbsenceType;
                            if (AttendanceCount.ContainsKey(periodName))
                            {
                                AttendanceCount[periodName]++;
                            }
                        }

                        System.Windows.Forms.ListViewItem.ListViewSubItem subitem = lvItem.SubItems[i];

                        foreach (K12.Data.AbsenceMappingInfo ai in _absenceList)
                        {
                            if (ai.Name != period.AbsenceType) continue;

                            subitem.Text = ai.Abbreviation;
                            break;
                        }
                    }
                }
            }
        }

        //統計資料
        private void BingDataMoralScore()
        {
            dgvAttendance.Rows.Clear();

            //明細統計
            dgvAttendance.Rows.Add(AttendanceDetailStatistics());

            //非明細統計
            dgvAttendance.Rows.Add(AttendanceInitialSummary());

            //AutoSummary
            dgvAttendance.Rows.Add(AttendanceAutoSummary());

            dgvAttendance.Columns[0].Width = 150;
        }

        private DataGridViewRow AttendanceDetailStatistics()
        {
            DataGridViewRow row1 = new DataGridViewRow();
            row1.CreateCells(dgvAttendance);
            row1.Cells[0].Value = "明細統計";
            row1.Cells[0].ReadOnly = true;
            row1.Cells[0].Style.BackColor = SetColor;

            //上色
            foreach (string each in AttendanceIndex.Keys)
            {
                if (AttendanceCount.ContainsKey(each))
                {
                    row1.Cells[AttendanceIndex[each]].Value = AttendanceCount[each] == 0 ? "" : AttendanceCount[each].ToString();
                }

                row1.Cells[AttendanceIndex[each]].Style.BackColor = SetColor;
                row1.Cells[AttendanceIndex[each]].ReadOnly = true;
            }

            return row1;
        }

        private DataGridViewRow AttendanceInitialSummary()
        {
            DataGridViewRow row2 = new DataGridViewRow();
            row2.CreateCells(dgvAttendance);
            row2.Cells[0].ReadOnly = true;
            row2.Cells[0].Style.BackColor = SetColor;
            row2.Cells[0].Value = "非明細統計";

            if (MSRecord != null)
            {
                SummarySetupObj obj = new SummarySetupObj(MSRecord.InitialSummary);

                foreach (AttendanceSetupObj objeach in obj.AttendanceList)
                {
                    if (AttendanceIndex.ContainsKey(objeach.PeritodTypeName))
                    {
                        row2.Cells[AttendanceIndex[objeach.PeritodTypeName]].Value = objeach.Count.ToString();
                    }
                }
            }

            return row2;

            //foreach (AutoSummaryRecord each in AutoRecord)
            //{
            //    if (each.SchoolYear == _SchoolYear && each.Semester == _Semester)
            //    {
            //        row2.CreateCells(dgvAttendance);
            //        row2.Cells[0].Value = "非明細統計";
            //        SummarySetupObj obj = new SummarySetupObj(each.InitialSummary);
            //        foreach (AttendanceSetupObj objeach in obj.AttendanceList)
            //        {
            //            if (AttendanceIndex.ContainsKey(objeach.PeritodTypeName))
            //            {
            //                row2.Cells[AttendanceIndex[objeach.PeritodTypeName]].Value = objeach.Count.ToString();
            //            }
            //        }
            //    }
            //}
        }

        private DataGridViewRow AttendanceAutoSummary()
        {
            DataGridViewRow row3 = new DataGridViewRow();
            row3.CreateCells(dgvAttendance);
            row3.Cells[0].Value = "本學期缺曠統計";
            row3.Cells[0].ReadOnly = true;
            row3.Cells[0].Style.BackColor = SetColor;

            //上色
            foreach (string each in AttendanceIndex.Keys)
            {
                row3.Cells[AttendanceIndex[each]].Style.BackColor = SetColor;
                row3.Cells[AttendanceIndex[each]].ReadOnly = true;
            }

            foreach (AutoSummaryRecord each in AutoRecord)
            {
                if (each.SchoolYear == _SchoolYear && each.Semester == _Semester)
                {
                    SummarySetupObj obj = new SummarySetupObj(each.AutoSummary);
                    foreach (AttendanceSetupObj objeach in obj.AttendanceList)
                    {
                        if (AttendanceIndex.ContainsKey(objeach.PeritodTypeName))
                        {
                            row3.Cells[AttendanceIndex[objeach.PeritodTypeName]].Value = objeach.Count.ToString();
                        }
                    }
                }
            }
            return row3;
        }

        //取得節次類型
        public static List<string> GetPeriodTypeItems()
        {
            List<string> list = new List<string>();
            List<K12.Data.PeriodMappingInfo> _periodList = K12.Data.PeriodMapping.SelectAll();

            foreach (K12.Data.PeriodMappingInfo each in _periodList)
            {
                if (!list.Contains(each.Type))
                    list.Add(each.Type);
            }

            list.Sort();

            return list;
        }

        //取得假別項目
        public static List<string> GetAbsenceItems()
        {
            List<string> list = new List<string>();
            List<K12.Data.AbsenceMappingInfo> _absenceList = K12.Data.AbsenceMapping.SelectAll();

            foreach (K12.Data.AbsenceMappingInfo each in _absenceList)
            {
                if (!list.Contains(each.Name))
                {
                    list.Add(each.Name);
                }
            }

            return list;
        }

        //新增
        private void btnAttendanceNew_Click(object sender, EventArgs e)
        {
            JHAttendanceRecord record = new JHAttendanceRecord();
            record.RefStudentID = _StudentID;

            AttendanceForm editor = new AttendanceForm(EditorStatus.Insert, record, _periodList, _UserPermission, _SchoolYear, _Semester);
            editor.ShowDialog();
        }

        //修改
        private void btnAttendanceEdit_Click(object sender, EventArgs e)
        {
            if (listViewAttendance.SelectedItems.Count == 1)
            {
                AttendanceForm editor = new AttendanceForm(EditorStatus.Update, listViewAttendance.SelectedItems[0].Tag as JHAttendanceRecord, _periodList, _UserPermission);
                editor.ShowDialog();
            }
        }

        //刪除
        private void btnAttendanceDelete_Click(object sender, EventArgs e)
        {
            if (listViewAttendance.SelectedItems.Count == 0)
            {
                FISCA.Presentation.Controls.MsgBox.Show("必須選擇一筆以上資料!!");
                return;
            }

            List<JHAttendanceRecord> AttendanceList = new List<JHAttendanceRecord>();

            if (FISCA.Presentation.Controls.MsgBox.Show("確定將刪除所選擇之缺曠資料?", "確認", MessageBoxButtons.YesNo) != DialogResult.Yes) return;

            foreach (ListViewItem item in listViewAttendance.SelectedItems)
            {
                JHAttendanceRecord record = item.Tag as JHAttendanceRecord;
                AttendanceList.Add(record);
            }     

            try
            {
                JHAttendance.Delete(AttendanceList);
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("刪除缺曠資料失敗" + ex.Message);
                return;
            }

            StringBuilder sb = new StringBuilder();
            sb.Append("學生「" + JHStudent.SelectByID(_StudentID).Name + "」");
            foreach (JHAttendanceRecord att in AttendanceList)
            {
                sb.AppendLine("日期「" + att.OccurDate.ToShortDateString() + "」");
            }
            sb.Append("缺曠資料已被刪除。");

            ApplicationLog.Log("學務系統.缺曠資料", "刪除學生缺曠資料", "student", _StudentID, sb.ToString());

            FISCA.Presentation.Controls.MsgBox.Show("刪除缺曠資料成功");
        }
        
        //儲存非明細統計值
        private void btnSaveAttendanceStatistics_Click(object sender, EventArgs e)
        {
            if (MSRecord != null) //是否有此物件
            {
                //log
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("修改學生「" + MSRecord.Student.Name + "」非明細缺曠統計值");

                SummarySetupObj obj = new SummarySetupObj(MSRecord.InitialSummary);
                obj.AttendanceList.Clear();

                foreach (DataGridViewCell cell in dgvAttendance.Rows[1].Cells)
                {
                    if ("" + cell.Value != "" && cell.OwningColumn.Index > 0)
                    {
                        string[] cellString = (cell.OwningColumn.Tag as string).Split(':');
                        AttendanceSetupObj attbigobj = new AttendanceSetupObj();
                        attbigobj.PeriodType = cellString[0];
                        attbigobj.Name = cellString[1];
                        attbigobj.Count = int.Parse("" + cell.Value);
                        //attbigobj.PeritodTypeName
                        obj.AttendanceList.Add(attbigobj);

                        //log
                        sb.AppendLine("類型「" + attbigobj.PeriodType + "」缺曠名稱「" + attbigobj.Name + "」統計值「" + attbigobj.Count + "」");
                    }
                }

                MSRecord.InitialSummary = obj.GetAllXmlElement();

                try
                {
                    JHMoralScore.Update(MSRecord);
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("更新,儲存失敗");
                    return;
                }

                //log
                ApplicationLog.Log("學務系統.缺曠學期統計", "修改學生非明細缺曠統計", "student", MSRecord.RefStudentID, sb.ToString());
            }
            else
            {
                //log
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("修改學生「" + JHStudent.SelectByID(_StudentID).Name + "」非明細缺曠統計值");

                JHMoralScoreRecord Jsr = new JHMoralScoreRecord();
                Jsr.RefStudentID = _StudentID;
                Jsr.SchoolYear = _SchoolYear;
                Jsr.Semester = _Semester;
                Jsr.InitialSummary = GetInitialSummary(dgvAttendance.Rows[1]);


                //log使用
                SummarySetupObj obj = new SummarySetupObj(Jsr.InitialSummary);
                foreach (AttendanceSetupObj each in obj.AttendanceList)
                {
                    sb.AppendLine("類型「" + each.PeriodType + "」缺曠名稱「" + each.Name + "」統計值「" + each.Count + "」");
                }

                try
                {
                    JHMoralScore.Insert(Jsr);
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("新增,儲存失敗");
                    return;
                }

                //log
                ApplicationLog.Log("學務系統.缺曠學期統計", "修改學生非明細缺曠統計", "student", _StudentID, sb.ToString());
            }

            lbHelp2.Text = "說明：白色欄位為可調整內容";
            lbHelp2.ForeColor = Color.Black;
            FISCA.Presentation.Controls.MsgBox.Show("儲存成功");
        }

        /// <summary>
        /// 由Row取得InitialSummary的Xml結構
        /// </summary>
        private XmlElement GetInitialSummary(DataGridViewRow row)
        {
            DSXmlHelper dsx = new DSXmlHelper("InitialSummary");
            dsx.AddElement("AttendanceStatistics");

            foreach (DataGridViewCell cell in dgvAttendance.Rows[1].Cells)
            {
                if ("" + cell.Value != "" && cell.OwningColumn.Index > 0)
                {
                    string[] cellString = (cell.OwningColumn.Tag as string).Split(':');

                    dsx.AddElement("AttendanceStatistics", "Absence");
                    dsx.SetAttribute("AttendanceStatistics/Absence", "Count", "" + cell.Value);
                    dsx.SetAttribute("AttendanceStatistics/Absence", "Name", cellString[1]);
                    dsx.SetAttribute("AttendanceStatistics/Absence", "PeriodType", cellString[0]);
                }

            }

            return dsx.BaseElement;
        }

        private int SchoolYearComparer(JHAttendanceRecord x, JHAttendanceRecord y)
        {
            return y.OccurDate.CompareTo(x.OccurDate);
        }

        private void btnExitAll_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dgvAttendance_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            btnSaveAttendanceStatistics.Pulse(20);
            lbHelp2.Text = "您已修改資料,請儲存缺曠統計";
            lbHelp2.ForeColor = Color.Red;
        }

        private void dgvAttendance_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dgvAttendance.EndEdit();

            ColumnAddInUnify.SetAddCell(dgvAttendance, true);

            dgvAttendance.BeginEdit(false);
        }

        private void listViewAttendance_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (listViewAttendance.SelectedItems.Count == 1)
            {
                AttendanceForm editor = new AttendanceForm(EditorStatus.Update, listViewAttendance.SelectedItems[0].Tag as JHAttendanceRecord, _periodList, _UserPermission);
                editor.ShowDialog();
            }
        }
    }
}
