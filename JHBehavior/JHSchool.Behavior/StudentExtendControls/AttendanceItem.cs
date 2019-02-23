using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Windows.Forms;
using FISCA.LogAgent;
using Framework;
using Framework.Security;
using JHSchool.Behavior.Feature;
using JHSchool.Behavior.Legacy;
using JHSchool.Data;
using JHSchool.Editor;
using FCode = Framework.Security.FeatureCodeAttribute;


namespace JHSchool.Behavior.StudentExtendControls
{
    //[FeatureCode("Content0060")]
    [FCode("JHSchool.Student.Detail0045", "缺曠記錄")]
    public partial class AttendanceItem : DetailContentBase
    {
        internal static FeatureAce UserPermission;

        private List<JHAttendanceRecord> _records = new List<JHAttendanceRecord>();
        private List<K12.Data.PeriodMappingInfo> _periodList = new List<K12.Data.PeriodMappingInfo>();
        private List<K12.Data.AbsenceMappingInfo> _absenceList = new List<K12.Data.AbsenceMappingInfo>();

        private BackgroundWorker BGW = new BackgroundWorker();
        private bool BkWBool = false;

        public AttendanceItem()
        {
            InitializeComponent();
            Group = "缺曠紀錄";

            JHAttendance.AfterDelete += new EventHandler<K12.Data.DataChangedEventArgs>(JHAttendance_AfterDelete);
            JHAttendance.AfterInsert += new EventHandler<K12.Data.DataChangedEventArgs>(JHAttendance_AfterDelete);
            JHAttendance.AfterUpdate += new EventHandler<K12.Data.DataChangedEventArgs>(JHAttendance_AfterDelete);

            //這是暫解法
            Attendance.Instance.ItemUpdated += new EventHandler<ItemUpdatedEventArgs>(Instance_ItemUpdated);

            //Initialize();

            //Attendance.Instance.ItemUpdated += new EventHandler<ItemUpdatedEventArgs>(Instance_ItemUpdated);

            BGW.DoWork += new DoWorkEventHandler(BkW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BkW_RunWorkerCompleted);

            UserPermission = User.Acl[FCode.GetCode(GetType())];

            btnAdd.Visible = UserPermission.Editable;
            btnUpdate.Visible = UserPermission.Editable;
            btnDelete.Visible = UserPermission.Editable;
            btnView.Visible = UserPermission.Viewable & !UserPermission.Editable;
        }

        void JHAttendance_AfterDelete(object sender, K12.Data.DataChangedEventArgs e)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<object, K12.Data.DataChangedEventArgs>(JHAttendance_AfterDelete), sender, e);
            }
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

        //這是暫解法
        void Instance_ItemUpdated(object sender, ItemUpdatedEventArgs e)
        {
            if (!BGW.IsBusy)
                BGW.RunWorkerAsync();
        }

        bool initialized = false;
        //建立預設畫面
        private void Initialize()
        {
            //取得此 Class 定義的 FeatureCode。
            //FeatureCodeAttribute code = Attribute.GetCustomAttribute(this.GetType(), typeof(FeatureCodeAttribute)) as FeatureCodeAttribute;
            //_permission = Framework.Legacy.GlobalOld.Acl[code.FeatureCode];
            if (initialized)
            {
                System.Threading.ThreadPool.QueueUserWorkItem(x =>
                {
                    _periodList = K12.Data.PeriodMapping.SelectAll();
                    _absenceList = K12.Data.AbsenceMapping.SelectAll();
                });
            }
            else
            {
                _periodList = K12.Data.PeriodMapping.SelectAll();
                _absenceList = K12.Data.AbsenceMapping.SelectAll();
            }

            listView.Columns.Clear();

            listView.Columns.Add("SchoolYear", "學年度");
            listView.Columns.Add("Semester", "學期");
            listView.Columns.Add("OccurDate", "缺曠日期");
            listView.Columns.Add("DayOfWeek", "星期");

            foreach (K12.Data.PeriodMappingInfo info in _periodList)
            {
                ColumnHeader column = listView.Columns.Add(info.Name, info.Name);
                column.Tag = info;
            }

            initialized = true;
        }

        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            //Attendance.Instance.SyncDataBackground(PrimaryKey);
            //base.OnPrimaryKeyChanged(e);

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

        void BkW_DoWork(object sender, DoWorkEventArgs e)
        {
            if (string.IsNullOrEmpty(this.PrimaryKey)) return;

            _records.Clear();
            _records = JHAttendance.SelectByStudentIDs(new string[] { this.PrimaryKey });
        }

        void BkW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (BkWBool)
            {
                BkWBool = false;
                BGW.RunWorkerAsync();
                return;
            }
            BindData();
            this.Loading = false;
        }

        private void BindData()
        {
            Initialize();

            listView.Items.Clear();

            _records.Sort(new Comparison<JHAttendanceRecord>(SchoolYearComparer));

            foreach (JHAttendanceRecord record in _records)
            {
                ListViewItem lvItem = listView.Items.Add(record.SchoolYear.ToString());

                lvItem.Tag = record;
                lvItem.SubItems.Add(record.Semester.ToString());
                lvItem.SubItems.Add(record.OccurDate.ToShortDateString());
                lvItem.SubItems.Add(record.DayOfWeek);
                if (listView.Columns.Count != 0)
                {
                    listView.Columns[0].Width = 70;
                    listView.Columns[1].Width = 50;
                    listView.Columns[2].Width = 110;
                }

                for (int i = 4; i < listView.Columns.Count; i++)
                    lvItem.SubItems.Add("");

                for (int i = 4; i < listView.Columns.Count; i++)
                {
                    ColumnHeader column = listView.Columns[i];
                    K12.Data.PeriodMappingInfo info = column.Tag as K12.Data.PeriodMappingInfo;

                    //if (record.PeriodDetail == null) continue;

                    foreach (K12.Data.AttendancePeriod period in record.PeriodDetail)
                    {
                        if (info == null) continue;
                        if (period.Period != info.Name) continue;
                        if (period.AbsenceType == null) continue;

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

        //private string chengDateTime(DateTime x)
        //{
        //    if (x == null)
        //        return "";
        //    string time = x.ToString();
        //    int y = time.IndexOf(' ');
        //    return time.Remove(y);
        //}

        private void btnAdd_Click(object sender, EventArgs e)
        {
            JHAttendanceRecord record = new JHAttendanceRecord();
            record.RefStudentID = this.PrimaryKey;

            AttendanceForm editor = new AttendanceForm(EditorStatus.Insert, record, _periodList, UserPermission);
            editor.ShowDialog();
          
            // (new SingleEditor(Student.Instance.SelectedList[0] )).ShowDialog();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (listView.SelectedItems.Count == 0)
            {
                MsgBox.Show("請先選擇一筆您要修改的資料");
                return;
            }
            else if (listView.SelectedItems.Count > 1)
            {
                MsgBox.Show("選擇資料筆數過多，一次只能修改一筆資料");
                return;
            }

            //(new SingleEditor(Student.Instance.SelectedList[0],(listView.SelectedItems[0].Tag as JHAttendanceRecord).OccurDate)).ShowDialog();
            AttendanceForm editor = new AttendanceForm(EditorStatus.Update, listView.SelectedItems[0].Tag as JHAttendanceRecord, _periodList, UserPermission);
            editor.ShowDialog();
        }
        //連點ListView時(因為使用JHSchool.Behavior.Legacy.ListViewEx 所以連點功能失效)
        //private void listView_MouseDoubleClick(object sender, MouseEventArgs e)
        //{
        //    btnUpdate_Click(null, EventArgs.Empty);
        //}

        private void btnView_Click(object sender, EventArgs e)
        {
            btnUpdate_Click(sender, e);
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (listView.SelectedItems.Count == 0)
            {
                MsgBox.Show("必須選擇一筆以上資料!!");
                return;
            }

            List<JHAttendanceRecord> AttendanceList = new List<JHAttendanceRecord>();

            if (MsgBox.Show("確定將刪除所選擇之缺曠資料?", "確認", MessageBoxButtons.YesNo) == DialogResult.No) return;

            foreach (ListViewItem item in listView.SelectedItems)
            {
                JHAttendanceRecord editor = item.Tag as JHAttendanceRecord;
                AttendanceList.Add(editor);
            }

            try
            {
                JHAttendance.Delete(AttendanceList);
            }
            catch (Exception ex)
            {
                MsgBox.Show("刪除缺曠資料失敗" + ex.Message);
                return;
            }

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("學生「" + JHStudent.SelectByID(this.PrimaryKey).Name + "」");
            foreach (JHAttendanceRecord att in AttendanceList)
            {
                sb.AppendLine("日期「" + att.OccurDate.ToShortDateString() + "」");
            }
            sb.AppendLine("缺曠資料已被刪除。");

            //DSXmlHelper LogDescription = new DSXmlHelper("Log");
            //LogDescription.AddElement("Description");
            //LogDescription.AddText("Description", sb.ToString());
            //LogDescription.AddElement("SubLogs");
            //LogDescription.AddElement("SubLogs", "Log");
            //LogDescription.SetAttribute("SubLogs/Log", "ID", editor.Student.ID);

            ApplicationLog.Log("學務系統.缺曠資料", "刪除學生缺曠資料", "student", this.PrimaryKey, sb.ToString());

            MsgBox.Show("刪除缺曠資料成功");
        }
        private int SchoolYearComparer(JHAttendanceRecord x, JHAttendanceRecord y)
        {
            return y.OccurDate.CompareTo(x.OccurDate);
        }

        private void listView_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (listView.SelectedItems.Count == 1)
            {

               // (new SingleEditor(Student.Instance.SelectedList[0], (listView.SelectedItems[0].Tag as JHAttendanceRecord).OccurDate)).ShowDialog();
                 AttendanceForm editor = new AttendanceForm(EditorStatus.Update, listView.SelectedItems[0].Tag as JHAttendanceRecord, _periodList, UserPermission);
                 editor.ShowDialog();
            }
        }

        private void listView_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
