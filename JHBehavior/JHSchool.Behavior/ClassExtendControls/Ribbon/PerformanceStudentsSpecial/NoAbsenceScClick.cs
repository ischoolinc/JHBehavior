using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JHSchool.Data;
using Aspose.Cells;
using JHSchool.Logic;
using K12.Data;
using DevComponents.DotNetBar.Controls;
using DevComponents.Editors;
using System.Drawing;
using System.ComponentModel;
using FISCA.Presentation.Controls;

namespace JHSchool.Behavior.ClassExtendControls.Ribbon
{
    class NoAbsenceScClick
    {
        PrintObj obj;

        BackgroundWorker BGW = new BackgroundWorker();

        //報表專用
        Workbook book;

        //自動統計內容
        List<AutoSummaryRecord> AutoSummaryList;

        //系統缺曠別
        public Dictionary<string, bool> AttendanceIsNoabsence = new Dictionary<string, bool>();

        bool _cbxSchoolYear1;
        int _intSchoolYear1;
        int _intSemester1;
        List<JHStudentRecord> _StudentRecordList;

        public void print(CheckBoxX cbxSchoolYear1,IntegerInput intSchoolYear1, IntegerInput intSemester1,List<JHStudentRecord> StudentRecordList)
        {
            _cbxSchoolYear1 = cbxSchoolYear1.Checked;
            _intSchoolYear1 = intSchoolYear1.Value;
            _intSemester1 = intSemester1.Value;
            _StudentRecordList = StudentRecordList;

            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            FISCA.Presentation.MotherForm.SetStatusBarMessage("全勤學生清單,列印中!");
            BGW.RunWorkerAsync();     
            
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            obj = new PrintObj();
            AutoSummaryList = new List<AutoSummaryRecord>();

            List<string> StudentNoAbsenceList = new List<string>(); //全勤的學生清單
            List<string> StudentAbsenceList = new List<string>(); //有缺曠的學生清單

            List<string> _StudentIDList = _StudentRecordList.Select(x => x.ID).ToList();
            AutoSummaryList.Clear();

            if (_cbxSchoolYear1)
            {
                //取得AutoSummary
                SchoolYearSemester SYS = new SchoolYearSemester(_intSchoolYear1, _intSemester1);
                AutoSummaryList = AutoSummary.Select(_StudentIDList, new SchoolYearSemester[] { SYS }, SummaryType.Attendance);
            }
            else
            {
                //取得AutoSummary
                AutoSummaryList = AutoSummary.Select(_StudentIDList, new List<SchoolYearSemester>(), SummaryType.Attendance);
            }

            #region 篩選出有記錄的學生(包含影響的缺曠)
            foreach (AutoSummaryRecord each in AutoSummaryList)
            {
                foreach (AbsenceCountRecord count in each.AbsenceCounts)
                {
                    if (count.Count == 0)
                        continue;
                    if (!AttendanceIsNoabsence.ContainsKey(count.Name)) //不包含假別中就離開
                        continue;
                    if (!AttendanceIsNoabsence[count.Name]) //True就是不影響全勤
                    {
                        if (!StudentAbsenceList.Contains(each.RefStudentID)) //如果沒有就加入
                        {
                            StudentAbsenceList.Add(each.RefStudentID);
                        }
                    }
                }
            }
            #endregion

            //排除有影響的學生就是我要的學生
            foreach (JHStudentRecord each in _StudentRecordList) //
            {
                if (!StudentAbsenceList.Contains(each.ID)) //不包含在有資料清單內
                {
                    StudentNoAbsenceList.Add(each.ID); //有全勤的學生
                }
            }

            book = new Workbook();
            Worksheet sheet = book.Worksheets[0];
            sheet.Name = "全勤學生清單";

            Cell A1 = sheet.Cells["A1"];
            A1.Style.Borders.SetColor(Color.Black);

            string A1Name = School.ChineseName + "　全勤學生清單";
            if (_cbxSchoolYear1)
            {
                A1Name += "　(" + _intSchoolYear1.ToString() + "/" + _intSemester1.ToString() + ")";
            }

            A1.PutValue(A1Name);
            A1.Style.HorizontalAlignment = TextAlignmentType.Center;
            sheet.Cells.Merge(0, 0, 1, 5);

            obj.FormatCell(sheet.Cells["A2"], "編號");
            obj.FormatCell(sheet.Cells["B2"], "班級");
            obj.FormatCell(sheet.Cells["C2"], "座號");
            obj.FormatCell(sheet.Cells["D2"], "姓名");
            obj.FormatCell(sheet.Cells["E2"], "學號");

            int index = 1;

            List<JHStudentRecord> studentList = JHStudent.SelectByIDs(StudentNoAbsenceList);

            studentList = SortClassIndex.JHSchoolData_JHStudentRecord(studentList);

            foreach (JHStudentRecord each in studentList)
            {
                int rowIndex = index + 2;
                obj.FormatCell(sheet.Cells["A" + rowIndex], index.ToString());
                obj.FormatCell(sheet.Cells["B" + rowIndex], each.Class.Name);
                obj.FormatCell(sheet.Cells["C" + rowIndex], each.SeatNo.HasValue ? each.SeatNo.Value.ToString() : "");
                obj.FormatCell(sheet.Cells["D" + rowIndex], each.Name);
                obj.FormatCell(sheet.Cells["E" + rowIndex], each.StudentNumber);
                index++;
            }
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                obj.PrintNow(book, "全勤學生清單");
                FISCA.Presentation.MotherForm.SetStatusBarMessage("列印全勤學生清單,已完成!");
            }
            else
            {
                MsgBox.Show("列印時發生錯誤!!" + e.Error.Message);
                FISCA.Presentation.MotherForm.SetStatusBarMessage("列印全勤學生清單,發生錯誤!");
            }
        }
    }
}
