using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using JHSchool.Behavior.BusinessLogic;
using JHSchool.Data;
using FISCA.DSAUtil;
using System.Xml;
using K12.Data.Configuration;
using FCode = Framework.Security.FeatureCodeAttribute;
using Framework.Security;
using Framework;
using JHSchool.Behavior.Feature;

namespace JHSchool.Behavior.StudentExtendControls
{
    [FCode("JHSchool.Student.Detail0037", "缺曠學期統計")]
    public partial class AttendanceUnifytIItem : DetailContentBase
    {
        private BackgroundWorker BGW = new BackgroundWorker();
        private bool BkWBool = false;
        private List<AutoSummaryRecord> AutoSummaryList = new List<AutoSummaryRecord>();

        private List<string> _periodTypes;
        private List<string> _absenceList;

        Dictionary<string, int> ColumnDic = new Dictionary<string, int>();

        internal static FeatureAce UserPermission;

        //建構子
        public AttendanceUnifytIItem()
        {

            InitializeComponent();

            BGW.DoWork += new DoWorkEventHandler(BgW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BgW_RunWorkerCompleted);

            JHAttendance.AfterDelete += new EventHandler<K12.Data.DataChangedEventArgs>(JHAttendance_AfterDelete);
            JHAttendance.AfterInsert += new EventHandler<K12.Data.DataChangedEventArgs>(JHAttendance_AfterDelete);
            JHAttendance.AfterUpdate += new EventHandler<K12.Data.DataChangedEventArgs>(JHAttendance_AfterDelete);

            JHMoralScore.AfterDelete += new EventHandler<K12.Data.DataChangedEventArgs>(JHAttendance_AfterDelete);
            JHMoralScore.AfterInsert += new EventHandler<K12.Data.DataChangedEventArgs>(JHAttendance_AfterDelete);
            JHMoralScore.AfterUpdate += new EventHandler<K12.Data.DataChangedEventArgs>(JHAttendance_AfterDelete);

            Group = "缺曠學期統計";

            UserPermission = User.Acl[FCode.GetCode(GetType())];
            if (!UserPermission.Editable)
            {
                this.listView.MouseDoubleClick -= new System.Windows.Forms.MouseEventHandler(this.listView_MouseDoubleClick);
                btnEdit.Enabled = UserPermission.Editable;
            }
        }

        //更新事件
        void JHAttendance_AfterDelete(object sender, K12.Data.DataChangedEventArgs e)
        {
            if (InvokeRequired)
                Invoke(new Action<object, K12.Data.DataChangedEventArgs>(JHAttendance_AfterDelete), sender, e);
            else
            {
                if (this.PrimaryKey != "")
                {
                    this.Loading = true;

                    if (BGW.IsBusy)
                    {
                        BkWBool = true;
                    }
                    else
                    {
                        BGW.RunWorkerAsync();
                    }
                }
            }
        }

        //切換學生
        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            if (this.PrimaryKey != "")
            {
                this.Loading = true;

                if (BGW.IsBusy)
                {
                    BkWBool = true;
                }
                else
                {
                    BGW.RunWorkerAsync();
                }
            }
        }

        //(舊)更新畫面
        void Instance_ItemUpdated(object sender, Framework.ItemUpdatedEventArgs e)
        {
            if (this.PrimaryKey != "")
            {
                this.Loading = true;

                if (BGW.IsBusy)
                {
                    BkWBool = true;
                }
                else
                {
                    BGW.RunWorkerAsync();
                }
            }
        }

        //背景
        void BgW_DoWork(object sender, DoWorkEventArgs e)
        {

            AutoSummaryList.Clear();
            AutoSummaryList = AutoSummary.Select(new string[] { this.PrimaryKey }, null);

            #region 駐解掉
            //List<SchoolYearSemester> SchoolSemesterList = new List<SchoolYearSemester>();

            //foreach (JHAttendanceRecord each in JHAttendance.SelectByStudentIDs(new string[] { this.PrimaryKey }))
            //{
            //    bool IsTrue = true;

            //    foreach (SchoolYearSemester school in SchoolSemesterList)
            //    {
            //        if (school.SchoolYear == each.SchoolYear && school.Semester == each.Semester)
            //        {
            //            IsTrue = false; //如果有重覆的
            //            break;
            //        }
            //    }

            //    if (IsTrue)
            //    {
            //        SchoolYearSemester SchoolSemester = new SchoolYearSemester(each.SchoolYear, each.Semester);
            //        SchoolSemesterList.Add(SchoolSemester);
            //    }
            //}

            //foreach (JHMoralScoreRecord each in JHMoralScore.SelectByStudentIDs(new string[] { this.PrimaryKey }))
            //{
            //    bool IsTrue = true;

            //    foreach (SchoolYearSemester school in SchoolSemesterList)
            //    {
            //        if (school.SchoolYear == each.SchoolYear && school.Semester == each.Semester)
            //        {
            //            IsTrue = false; //如果有重覆的
            //            break;
            //        }
            //    }

            //    if (IsTrue)
            //    {
            //        SchoolYearSemester SchoolSemester = new SchoolYearSemester(each.SchoolYear, each.Semester);
            //        SchoolSemesterList.Add(SchoolSemester);
            //    }
            //} 
            #endregion
        }

        //背景完成
        void BgW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (BkWBool)
            {
                BkWBool = false;
                BGW.RunWorkerAsync();
                return;
            }

            BindColumn();

            BindData();

            this.Loading = false;
        }

        //將節次類別,填入欄位
        private void BindColumn()
        {
            listView.Columns.Clear();
            ColumnDic.Clear();

            ColumnHeader chSchoolYear = new ColumnHeader();
            chSchoolYear.Text = "學年度";
            int CHint = listView.Columns.Add(chSchoolYear);
            listView.Columns[CHint].Width = (3 - 1) * 13 + 31;
            ColumnDic.Add("學年度", CHint);

            ColumnHeader chSemester = new ColumnHeader();
            chSemester.Text = "學期";
            CHint = listView.Columns.Add(chSemester);
            listView.Columns[CHint].Width = (2 - 1) * 13 + 31;
            ColumnDic.Add("學期", CHint);

            _periodTypes = GetPeriodTypeItems();
            _periodTypes.Sort();
            _absenceList = GetAbsenceItems();

            foreach (string periodType in _periodTypes)
            {
                foreach (string each in _absenceList)
                {
                    string columnName = periodType + each;
                    ColumnHeader cHeader = new ColumnHeader();
                    cHeader.Text = columnName;
                    cHeader.Tag = periodType + ":" + each;
                    CHint = listView.Columns.Add(cHeader);
                    listView.Columns[CHint].Width = (columnName.Length - 1) * 13 + 31;
                    ColumnDic.Add(columnName, CHint);
                }
            }
        }

        //更新畫面資料
        private void BindData()
        {
            if (!this.SaveButtonVisible && !this.CancelButtonVisible && this.PrimaryKey.Contains(PrimaryKey))
            {
                this.listView.Items.Clear();

                AutoSummaryList.Sort(new Comparison<AutoSummaryRecord>(SortAutoSummary));

                foreach (AutoSummaryRecord each in AutoSummaryList)
                {
                    SummarySetupObj obj = new SummarySetupObj(each.AutoSummary);

                    ListViewItem itms = new ListViewItem(each.SchoolYear.ToString());
                    itms.SubItems.Add(each.Semester.ToString());

                    for (int x = 0; x < ColumnDic.Count; x++)
                    {
                        itms.SubItems.Add("");
                    }

                    foreach (AttendanceSetupObj AttEach in obj.AttendanceList)
                    {
                        if (ColumnDic.ContainsKey(AttEach.PeritodTypeName))
                        {
                            itms.SubItems[ColumnDic[AttEach.PeritodTypeName]].Text = AttEach.Count.ToString();
                        }
                    }


                    itms.Tag = each;
                    listView.Items.Add(itms);
                }
            }
        }

        //取得節次類型
        public static List<string> GetPeriodTypeItems()
        {
            List<string> list = new List<string>();
            List<PeriodMappingInfo> _periodList = QueryPeriodMapping.Load();

            foreach (PeriodMappingInfo each in _periodList)
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
            List<AbsenceMappingInfo> _absenceList = QueryAbsenceMapping.Load();

            foreach (AbsenceMappingInfo each in _absenceList)
            {
                if (!list.Contains(each.Name))
                {
                    list.Add(each.Name);
                }
            }

            return list;
        }

        //學期判斷
        private int SortAutoSummary(AutoSummaryRecord x, AutoSummaryRecord y)
        {
            string SchoolYearSemester1 = x.SchoolYear.ToString().PadLeft(3, '0') + x.Semester.ToString().PadLeft(3, '0');
            string SchoolYearSemester2 = y.SchoolYear.ToString().PadLeft(3, '0') + y.Semester.ToString().PadLeft(3, '0');

            return SchoolYearSemester1.CompareTo(SchoolYearSemester2);
        }

        //開啟編輯畫面
        private void listView_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (listView.SelectedItems.Count == 1)
            {
                AutoSummaryRecord auto = (AutoSummaryRecord)listView.SelectedItems[0].Tag;

                AttendanceUnifyForm DemeritForm = new AttendanceUnifyForm(auto.RefStudentID, auto.SchoolYear, auto.Semester, UserPermission);
                DemeritForm.ShowDialog();
                if (!BGW.IsBusy)
                {
                    BGW.RunWorkerAsync();
                }
            }
        }

        //開啟編輯畫面
        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (listView.SelectedItems.Count == 1)
            {
                AutoSummaryRecord auto = (AutoSummaryRecord)listView.SelectedItems[0].Tag;

                AttendanceUnifyForm AttendanceForm = new AttendanceUnifyForm(auto.RefStudentID, auto.SchoolYear, auto.Semester, UserPermission);
                AttendanceForm.ShowDialog();
                if (!BGW.IsBusy)
                {
                    BGW.RunWorkerAsync();
                }
            }
        }
    }
}
