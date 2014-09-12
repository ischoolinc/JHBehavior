using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using JHSchool.Behavior.StuAdminExtendControls;
using FISCA.DSAUtil;
using System.Xml;
using JHSchool.Data;
using Aspose.Cells;
using System.IO;
using System.Diagnostics;
using JHSchool.Logic;
using K12.Data;

namespace JHSchool.Behavior.ClassExtendControls.Ribbon
{
    public partial class SpecialForm : BaseForm
    {
        PrintObj obj = new PrintObj();

        BackgroundWorker BGW = new BackgroundWorker();

        List<string> AttendanceStringList = new List<string>();

        Dictionary<string, bool> AttendanceIsNoabsence = new Dictionary<string, bool>();

        //學生清單
        List<JHStudentRecord> _StudentRecordList = new List<JHStudentRecord>();

        public SpecialForm()
        {
            InitializeComponent();
        }

        //載入預設畫面
        private void SpecialForm_Load(object sender, EventArgs e)
        {
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);

            tabControl1.Enabled = false;
            this.Text = "學生資料讀取中";
            BGW.RunWorkerAsync();
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            _StudentRecordList = obj.GetStudentList();
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Text = "查詢學生特殊表現名單";
            tabControl1.Enabled = true;
            SetSchoolYearSemester();

            SetForm1();
        }

        //預設畫面1內容(累計節次,假別)
        private void SetForm1()
        {
            txtPeriodCount.Text = "1";

            //缺曠別
            AttendanceStringList.Clear();
            AttendanceIsNoabsence.Clear();
            listViewEx1.Items.Clear();

            List<K12.Data.AbsenceMappingInfo> InfoList = K12.Data.AbsenceMapping.SelectAll();
            foreach (K12.Data.AbsenceMappingInfo e in InfoList)
            {
                AttendanceStringList.Add(e.Name);
                listViewEx1.Items.Add(e.Name);

                AttendanceIsNoabsence.Add(e.Name, e.Noabsence);
            }
        }

        //列印"缺曠累計名單"
        private void btnPrint1_Click(object sender, EventArgs e)
        {
            AttendanceScClick Atsc = new AttendanceScClick();
            Atsc.AttendanceStringList = AttendanceStringList;
            Atsc.print(cbxSchoolYear1, intSchoolYear1, intSemester1, _StudentRecordList, txtPeriodCount, listViewEx1);
        }

        //列印"全勤學生"名單
        private void btnPrint2_Click(object sender, EventArgs e)
        {
            NoAbsenceScClick NAsc = new NoAbsenceScClick();
            NAsc.AttendanceIsNoabsence = AttendanceIsNoabsence;
            NAsc.print(cbxSchoolYear1, intSchoolYear1, intSemester1, _StudentRecordList);
        }

        //列印"獎勵特殊表現"名單
        private void btnPrint4_Click(object sender, EventArgs e)
        {
            MeritScClick Msc = new MeritScClick();
            Msc.print(cbxSchoolYear1, intSchoolYear1, intSemester1, _StudentRecordList, tbMeritA, tbMeritB, tbMeritC, cbxIgnoreDemerit, cbxDemeritIsNull, cbxIsDemeritClear);
        }

        //列印"懲戒特殊表現"名單
        private void btnPrint3_Click(object sender, EventArgs e)
        {
            DemeritScClick Dmsc = new DemeritScClick();
            Dmsc.Print(cbxSchoolYear1, intSchoolYear1, intSemester1, _StudentRecordList, tbDemeritA, tbDemeritB, tbDemeritC, cbxIsMeritAndDemerit);
        }

        //預設畫面的學年度學期
        private void SetSchoolYearSemester()
        {
            int SchoolYear;
            int Semester;

            if (int.TryParse(School.DefaultSchoolYear, out SchoolYear))
            {
                intSchoolYear1.Value = SchoolYear;
            }

            if (int.TryParse(School.DefaultSemester, out Semester))
            {
                intSemester1.Value = Semester;
            }

        }

        //全選假別內容
        private void cbxSelectAllPeriod_CheckedChanged(object sender, EventArgs e)
        {
            foreach (ListViewItem each in listViewEx1.Items)
            {
                each.Checked = cbxSelectAllPeriod.Checked;
            }
        }

        //離開本功能
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #region Link

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            AbsenceConfigForm config = new AbsenceConfigForm();
            config.ShowDialog();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ReduceForm config = new ReduceForm();
            config.ShowDialog();
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ReduceForm config = new ReduceForm();
            config.ShowDialog();
        }
        #endregion

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            BalanceConfigForm BCForm = new BalanceConfigForm();
            BCForm.ShowDialog();
        }
    }
}
