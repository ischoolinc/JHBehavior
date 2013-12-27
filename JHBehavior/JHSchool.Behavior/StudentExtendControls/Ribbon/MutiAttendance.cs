using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using FISCA.DSAUtil;
using System.Xml;
using JHSchool.Behavior.StuAdminExtendControls;
using JHSchool.Data;
using JHSchool.Behavior.Legacy.AttendanceEditor;

namespace JHSchool.Behavior.StudentExtendControls.Ribbon
{
    public partial class MutiAttendance : BaseForm
    {
        List<JHStudentRecord> StudentList = new List<JHStudentRecord>();
        List<JHAttendanceRecord> AttendanceList = new List<JHAttendanceRecord>();

        Dictionary<string, MutiStudentObj> dic = new Dictionary<string, MutiStudentObj>();

        BackgroundWorker BGWLoad = new BackgroundWorker();
        BackgroundWorker BGWSave = new BackgroundWorker();

        DateTime _dtiEndTime = new DateTime();
        DateTime _dtiStartTime = new DateTime();

        public MutiAttendance()
        {
            InitializeComponent();
        }

        private void MutiAttendance_Load(object sender, EventArgs e)
        {
            BGWLoad.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGWLoad.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);

            BGWSave.DoWork += new DoWorkEventHandler(BGWSave_DoWork);
            BGWSave.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGWSave_RunWorkerCompleted);

            //日期區間預設一週
            dtiStartTime.Value = DateTime.Now;
            dtiEndTime.Value = DateTime.Now.AddDays(7);

            //開始取得資料
            btnSave.Enabled = false;
            BGWLoad.RunWorkerAsync();
        }

        //取資料背景模式
        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            //取得學生,並進行排序
            StudentList.Clear();
            StudentList = JHStudent.SelectByIDs(K12.Presentation.NLDPanels.Student.SelectedSource);
            StudentList.Sort(new Comparison<JHStudentRecord>(SortStudentInAttendance));

            //取得學生缺曠資料
            AttendanceList.Clear();
            AttendanceList = JHAttendance.SelectByStudentIDs(K12.Presentation.NLDPanels.Student.SelectedSource);
        }

        //完成建立畫面
        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //取得選取學生
            SetStudentList();

            //取得缺曠別
            SetAbsence();

            //取得節次別
            SetPeriod();

            btnSave.Enabled = true;
        }

        /// <summary>
        /// 建立畫面學生清單
        /// </summary>
        private void SetStudentList()
        {
            dic.Clear();

            foreach (JHStudentRecord each in StudentList)
            {
                //建立畫面資料
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridViewX1);
                row.Cells[0].Value = each.Class != null ? each.Class.Name : "";
                row.Cells[1].Value = each.SeatNo.HasValue ? each.SeatNo.Value.ToString() : "";
                row.Cells[2].Value = each.Name;
                dataGridViewX1.Rows.Add(row);

                //建立物件清單
                MutiStudentObj obj = new MutiStudentObj(each.ID);
                dic.Add(each.ID, obj);
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            _dtiEndTime = new DateTime(dtiEndTime.Value.Year,dtiEndTime.Value.Month,dtiEndTime.Value.Day);
            _dtiStartTime = new DateTime(dtiStartTime.Value.Year, dtiStartTime.Value.Month, dtiStartTime.Value.Day);
            BGWSave.RunWorkerAsync();

        }

        void BGWSave_DoWork(object sender, DoWorkEventArgs e)
        {
            //日期區間確認(包含設定值所排除之星期次)
            List<DateTime> DateTimeList = GetDateTimeRange();

            //把缺曠清單建立於學生物件內
            foreach (JHAttendanceRecord each in AttendanceList)
            {
                if (dic.ContainsKey(each.RefStudentID)) //有此物件
                {
                    if (DateTimeList.Contains(each.OccurDate)) //於日期區間內
                    {
                        dic[each.RefStudentID].AttendList.Add(each);
                    }
                }
            }

            //建立新增或是更新清單

            foreach (DateTime each in DateTimeList)
            {
                foreach (string student in dic.Keys)
                {
                    dic[student].SetupAttendance(each);
                }
            }


            if (rbCoverOverFalse.Checked) //(預設)略過原有假別
            {

            }
            else //覆蓋原有假別
            {

            }



        }

        void BGWSave_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MsgBox.Show("資料儲存成功!");
        }


        /// <summary>
        /// 取得日期區間清單
        /// </summary>
        /// <returns></returns>
        private List<DateTime> GetDateTimeRange()
        {
            //取得星期設定清單
            List<string> WeekStringList = GetWeekList();
            //取得星期對照表
            List<DayOfWeek> WeekList = ChengDayOfWeel(WeekStringList);

            List<DateTime> DateTimeList = new List<DateTime>();

            for (DateTime dt = _dtiStartTime; dt <= _dtiEndTime; )
            {
                //判斷是否已存在

                //排除已存在 and 星期設定之內容
                if (!DateTimeList.Contains(dt) && WeekList.Contains(dt.DayOfWeek))
                {

                    DateTimeList.Add(dt);
                }

                //是否與星期設定相衝突


                dt = dt.AddDays(1);
            }

            return DateTimeList;
        }

        /// <summary>
        /// 取得星期對照表
        /// </summary>
        private List<DayOfWeek> ChengDayOfWeel(List<string> list)
        {
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
        }

        /// <summary>
        /// 取得設定星期清單
        /// </summary>
        private List<string> GetWeekList()
        {
            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["缺曠批次登錄_星期設定_多人多天缺曠登錄"];
            string cdIN = cd["星期設定"];

            DSXmlHelper dsx = new DSXmlHelper("WeekSetup");

            XmlElement day;

            if (cdIN != "")
            {
                day = DSXmlHelper.LoadXml(cdIN);
            }
            else
            {
                day = dsx.BaseElement;
            }

            List<string> list = new List<string>();
            foreach (XmlElement each in day.SelectNodes("Day"))
            {
                list.Add(each.GetAttribute("Detail"));
            }

            return list;
        }

        /// <summary>
        /// 設定缺曠類別
        /// </summary>
        private void SetAbsence()
        {
            DSResponse dsrsp = JHSchool.Feature.Legacy.Config.GetAbsenceList();
            DSXmlHelper helper = dsrsp.GetContent();
            foreach (XmlElement element in helper.GetElements("Absence"))
            {
                //建立缺曠物件
                AbsenceInfo info = new AbsenceInfo(element);
                cbxAbsenceList.Items.Add(info.Name);
            }

            cbxAbsenceList.SelectedIndex = 0;
        }

        /// <summary>
        /// 設定每日節次
        /// </summary>
        private void SetPeriod()
        {

            DSResponse dsrsp = JHSchool.Feature.Legacy.Config.GetPeriodList();
            DSXmlHelper helper = dsrsp.GetContent();
            List<PeriodInfo> collection = new List<PeriodInfo>();
            foreach (XmlElement element in helper.GetElements("Period"))
            {
                PeriodInfo info = new PeriodInfo(element);

                DataGridViewCheckBoxColumn colCheck = new DataGridViewCheckBoxColumn();
                colCheck.Name = info.Name;
                colCheck.HeaderText = info.Name;
                dgvPeriodType.Columns.Add(colCheck);
            }

            dgvPeriodType.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.ColumnHeader);
            dgvPeriodType.Rows.Add();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 學生排序功能
        /// </summary>
        private int SortStudentInAttendance(JHStudentRecord xStud, JHStudentRecord yStud)
        {
            string xClass = xStud.Class != null ? xStud.Class.Name : "";
            xClass = xClass.PadLeft(6, '0');
            string yClass = yStud.Class != null ? yStud.Class.Name : "";
            yClass = yClass.PadLeft(6, '0');

            string xSean = xStud.SeatNo.HasValue ? xStud.SeatNo.Value.ToString() : "";
            xSean = xSean.PadLeft(6, '0');
            string ySean = yStud.SeatNo.HasValue ? yStud.SeatNo.Value.ToString() : "";
            ySean = ySean.PadLeft(6, '0');

            string xx = xClass + xSean;
            string yy = yClass + ySean;
            return xx.CompareTo(yy);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Searchday Sday = new Searchday("缺曠批次登錄_星期設定_多人多天缺曠登錄");
            Sday.ShowDialog();
        }
    }
}
