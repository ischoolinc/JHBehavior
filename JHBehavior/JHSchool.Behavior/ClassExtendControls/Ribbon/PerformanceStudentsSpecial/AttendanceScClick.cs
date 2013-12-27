using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JHSchool.Data;
using JHSchool.Logic;
using K12.Data;
using DevComponents.DotNetBar;
using DevComponents.Editors;
using DevComponents.DotNetBar.Controls;
using FISCA.Presentation.Controls;
using System.Windows.Forms;
using Aspose.Cells;
using System.ComponentModel;

namespace JHSchool.Behavior.ClassExtendControls.Ribbon
{
    class AttendanceScClick
    {
        //自動統計內容
        List<AutoSummaryRecord> AutoSummaryList;

        BackgroundWorker BGW = new BackgroundWorker();

        //
        PrintObj obj;

        //系統的缺曠別清單
        public List<string> AttendanceStringList = new List<string>();

        //報表專用
        Workbook book;

        bool _cbxSchoolYear1;
        int _intSchoolYear1;
        int _intSemester1;
        List<JHStudentRecord> _StudentRecordList;
        string _txtPeriodCount;

        //缺曠清單
        List<string> SelectAbsenceList = new List<string>();

        public void print(CheckBoxX cbxSchoolYear1, IntegerInput intSchoolYear1, IntegerInput intSemester1, List<JHStudentRecord> StudentRecordList, TextBoxX txtPeriodCount, ListViewEx listViewEx1)
        {
            _cbxSchoolYear1 = cbxSchoolYear1.Checked;
            _intSchoolYear1 = intSchoolYear1.Value;
            _intSemester1 = intSemester1.Value;
            _StudentRecordList = StudentRecordList;
            _txtPeriodCount = txtPeriodCount.Text;

            //選擇的缺曠類別
            SelectAbsenceList.Clear();
            foreach (ListViewItem item in listViewEx1.Items)
            {
                if (item.Checked)
                {
                    SelectAbsenceList.Add(item.Text);
                }
            }

            if (SelectAbsenceList.Count == 0)
            {
                MsgBox.Show("至少必須選擇一個缺曠類別!");
                return;
            }

            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
            FISCA.Presentation.MotherForm.SetStatusBarMessage("缺曠累計名單,列印中!");
            BGW.RunWorkerAsync();
        }
        //開始背景模式
        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            obj = new PrintObj();
            AutoSummaryList = new List<AutoSummaryRecord>();
            //累計節次
            int PeriodCount;

            //取得缺曠內容
            //List<JHAttendanceRecord> AttendanceList;

            List<string> _StudentIDList = _StudentRecordList.Select(x => x.ID).ToList();
            AutoSummaryList.Clear();

            if (_cbxSchoolYear1)
            {
                //AttendanceList = JHAttendance.SelectBySchoolYearAndSemester(_StudentRecordList, intSchoolYear1.Value, intSemester1.Value);

                //取得AutoSummary
                SchoolYearSemester SYS = new SchoolYearSemester(_intSchoolYear1, _intSemester1);
                AutoSummaryList = AutoSummary.Select(_StudentIDList, new SchoolYearSemester[] { SYS }, SummaryType.Attendance);
            }
            else
            {
                //AttendanceList = JHAttendance.SelectByStudentIDs(_StudentIDList);

                //取得AutoSummary
                AutoSummaryList = AutoSummary.Select(_StudentIDList, new List<SchoolYearSemester>(), SummaryType.Attendance);
            }

            if (!int.TryParse(_txtPeriodCount, out PeriodCount))
            {
                MsgBox.Show("累計節次內容非數字!!");
                return;
            }

            //缺曠總數量
            //Dictionary<string, Dictionary<string, int>> studentAttendance1 = new Dictionary<string, Dictionary<string, int>>();

            //缺曠權重比
            Dictionary<string, Dictionary<string, double>> studentAttendance2 = new Dictionary<string, Dictionary<string, double>>();

            //AttendanceList.Sort(new SortClass().SortAttendanceRecord);
            //AutoSummaryList.Sort(new SortClass().SortAutoSummaryRecord);

            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["特殊學生表現_缺曠累積名單"];


            #region 建立目前資料字典
            foreach (AutoSummaryRecord auto in AutoSummaryList)
            {
                //字典是否有此學生
                if (!studentAttendance2.ContainsKey(auto.RefStudentID))
                {
                    //studentAttendance1.Add(attendance.RefStudentID, new Dictionary<string, int>());
                    studentAttendance2.Add(auto.RefStudentID, new Dictionary<string, double>());
                }
                foreach (AbsenceCountRecord absence in auto.AbsenceCounts)
                {
                    //如果不是系統中假別及節次,不予計算
                    if (!AttendanceStringList.Contains(absence.Name)/* || !_PeriodDouble.ContainsKey(attenancePeriod.Period)*/)
                        continue;
                    //此學生字典是否已有此假別
                    if (!studentAttendance2[auto.RefStudentID].ContainsKey(absence.Name))
                    {
                        //studentAttendance1[attendance.RefStudentID].Add(attenancePeriod.AbsenceType, 0); //預設為0
                        studentAttendance2[auto.RefStudentID].Add(absence.Name, 0); //預設為0
                    }
                    //記數
                    //studentAttendance1[attendance.RefStudentID][attenancePeriod.AbsenceType]++;

                    //記權重(進行換算)
                    if (cd.Contains(absence.PeriodType))
                    {
                        if (obj.doubleCheck(cd[absence.PeriodType])) //如果是double就乘上基數
                        {
                            studentAttendance2[auto.RefStudentID][absence.Name] += absence.Count * double.Parse(cd[absence.PeriodType]);
                        }
                        else //不是就預設為1
                        {
                            studentAttendance2[auto.RefStudentID][absence.Name] += absence.Count * 1;
                        }
                    }
                    else
                    {
                        studentAttendance2[auto.RefStudentID][absence.Name] += absence.Count * 1;
                    }
                }
            }
            #endregion

            #region 預設報表資訊
            book = new Workbook();
            book.Worksheets.Clear();

            int sheetIndex = book.Worksheets.Add();
            Worksheet sheet = book.Worksheets[sheetIndex];
            sheet.Name = "缺曠累計名單";

            //將格子合併
            sheet.Cells.Merge(0, 0, 1, 5 + SelectAbsenceList.Count);
            sheet.Cells[0, 0].PutValue(School.ChineseName + "　缺曠累計名單");
            sheet.Cells[0, 0].Style.HorizontalAlignment = TextAlignmentType.Center;
            sheet.Cells[1, 0].PutValue("班級");
            sheet.Cells[1, 1].PutValue("座號");
            sheet.Cells[1, 2].PutValue("姓名");
            sheet.Cells[1, 3].PutValue("學號");

            Dictionary<string, int> saveAttAddress1 = new Dictionary<string, int>();

            int countList = 4;
            foreach (string var in SelectAbsenceList) //依選擇的假別
            {
                saveAttAddress1.Add(var, countList); //記錄定位
                sheet.Cells[1, countList].PutValue(var);
                countList++;
            }
            sheet.Cells[1, countList].PutValue("累積節次");

            int cellcount = 2; //ROW的Index
            //int _MergeInt = 0; 
            #endregion


            List<JHStudentRecord> PrintStudentList = JHStudent.SelectByIDs(studentAttendance2.Keys);
            PrintStudentList = SortClassIndex.JHSchoolData_JHStudentRecord(PrintStudentList);

            #region 學生逐一列印
            //取得一名學生之資料
            foreach (JHStudentRecord student in PrintStudentList)
            {
                string var = student.ID;

                double xyz = 0;
                //處理假別相加
                foreach (string invar in studentAttendance2[var].Keys)
                {
                    //假別是否是使用者所選
                    if (SelectAbsenceList.Contains(invar))
                    {
                        //將資料相加
                        xyz = xyz + studentAttendance2[var][invar];
                    }
                }

                //如果累計數量大於等於使用者所輸入
                if (xyz >= float.Parse(_txtPeriodCount))
                {
                    //班級
                    sheet.Cells[cellcount, 0].PutValue(student.Class.Name);
                    //座號
                    sheet.Cells[cellcount, 1].PutValue(student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "");
                    //姓名
                    sheet.Cells[cellcount, 2].PutValue(student.Name);
                    //學號
                    sheet.Cells[cellcount, 3].PutValue(student.StudentNumber);
                    //權重累積
                    sheet.Cells[cellcount, countList].PutValue(xyz);

                    //數量
                    foreach (string invar in studentAttendance2[var].Keys) //取得假別
                    {
                        if (saveAttAddress1.ContainsKey(invar)) //是否有在定位內
                        {
                            sheet.Cells[cellcount, saveAttAddress1[invar]].PutValue(studentAttendance2[var][invar]);
                        }
                    }

                    //補0機制...
                    foreach (int injar in saveAttAddress1.Values)
                    {
                        if (sheet.Cells[cellcount, injar].StringValue == string.Empty)
                        {
                            sheet.Cells[cellcount, injar].PutValue(0);
                        }
                        //_MergeInt = injar;
                    }


                    cellcount++;
                }
            }
            #endregion
        }

        //列印完成
        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                obj.PrintNow(book, "缺曠累計名單");
                FISCA.Presentation.MotherForm.SetStatusBarMessage("列印缺曠累計名單,已完成!");
            }
            else
            {
                MsgBox.Show("列印時發生錯誤!!" + e.Error.Message);
                FISCA.Presentation.MotherForm.SetStatusBarMessage("列印缺曠累計名單,發生錯誤!");

            }
        }
    }
}
