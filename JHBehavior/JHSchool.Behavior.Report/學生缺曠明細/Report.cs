using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Aspose.Cells;
using FISCA.DSAUtil;
using Framework;
using JHSchool.Data;
using K12.Data;
using K12.Logic;

namespace JHSchool.Behavior.Report.學生缺曠明細
{
    internal class Report : IReport
    {
        #region IReport 成員

        private BackgroundWorker _BGWAbsenceDetail;
        SelectAttendanceForm form;

        public void Print()
        {
            if (Student.Instance.SelectedList.Count == 0)
                return;

            //警告使用者別做傻事
            if (Student.Instance.SelectedList.Count > 1500)
            {
                MsgBox.Show("您選取的學生超過 1500 個，可能會發生意想不到的錯誤，請減少選取的學生。");
                return;
            }

            form = new SelectAttendanceForm();

            if (form.ShowDialog() == DialogResult.OK)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化學生個人缺曠明細...");

                //object[] args = new object[] { form.SchoolYear, form.Semester };

                _BGWAbsenceDetail = new BackgroundWorker();
                _BGWAbsenceDetail.DoWork += new DoWorkEventHandler(_BGWAbsenceDetail_DoWork);
                _BGWAbsenceDetail.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CommonMethods.ExcelReport_RunWorkerCompleted);
                _BGWAbsenceDetail.ProgressChanged += new ProgressChangedEventHandler(CommonMethods.Report_ProgressChanged);
                _BGWAbsenceDetail.WorkerReportsProgress = true;
                _BGWAbsenceDetail.RunWorkerAsync();
            }
        }

        #endregion

        void _BGWAbsenceDetail_DoWork(object sender, DoWorkEventArgs e)
        {
            string reportName = "學生個人缺曠明細";

            #region 快取相關資料

            //選擇的學生
            List<JHStudentRecord> selectedStudents = JHStudent.SelectByIDs(Student.Instance.SelectedKeys);
            selectedStudents = SortClassIndex.JHSchoolData_JHStudentRecord(selectedStudents);

            //紀錄所有學生ID
            List<string> allStudentID = new List<string>();

            //對照表
            Dictionary<string, string> absenceList = new Dictionary<string, string>();
            Dictionary<string, string> periodList = new Dictionary<string, string>();

            //根據節次類型統計每一種缺曠別的次數
            Dictionary<string, Dictionary<string, Dictionary<string, int>>> periodStatisticsByType = new Dictionary<string, Dictionary<string, Dictionary<string, int>>>();

            //每一位學生的缺曠明細
            Dictionary<string, Dictionary<string, Dictionary<string, string>>> studentAbsenceDetail = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();

            //每一位學生ID,會有多筆的AutoSummary
            Dictionary<string, List<AutoSummaryRecord>> AutoSummaryDic = new Dictionary<string, List<AutoSummaryRecord>>();

            //紀錄每一個節次在報表中的 column index
            Dictionary<string, int> columnTable = new Dictionary<string, int>();

            //取得所有學生ID
            foreach (JHStudentRecord var in selectedStudents)
            {
                allStudentID.Add(var.ID);
            }

            //取得 Absence List
            List<K12.Data.AbsenceMappingInfo> AbsenceList = K12.Data.AbsenceMapping.SelectAll();
            foreach (K12.Data.AbsenceMappingInfo var in AbsenceList)
            {
                string name = var.Name;

                if (!absenceList.ContainsKey(name))
                    absenceList.Add(name, var.Abbreviation);
            }

            //取得 Period List
            List<K12.Data.PeriodMappingInfo> PeriodList = K12.Data.PeriodMapping.SelectAll();
            PeriodList.Sort(tool.SortPeriod);

            foreach (K12.Data.PeriodMappingInfo var in PeriodList)
            {
                string name = var.Name;
                string type = var.Type;

                if (!periodList.ContainsKey(name))
                    periodList.Add(name, type);
            }

            //產生 DSRequest，取得缺曠明細
            //取得獎懲資料
            List<AttendanceRecord> AttendanceList = new List<AttendanceRecord>();

            //取得自動統計資料
            List<AutoSummaryRecord> AutoSummaryList = new List<AutoSummaryRecord>();

            if (form.SelectDayOrSchoolYear) //依日期
            {
                AttendanceList = Attendance.Select(allStudentID, form.StartDay, form.EndDay, null, null, null);
            }
            else
            {
                if (form.checkBoxX1Bool) //全部學期列印
                {
                    AutoSummaryList = AutoSummary.Select(allStudentID, null);
                    AttendanceList = Attendance.SelectByStudentIDs(allStudentID);
                }
                else
                {
                    SchoolYearSemester SYS = new SchoolYearSemester(int.Parse(form.SchoolYear), int.Parse(form.Semester));
                    List<SchoolYearSemester> SYSList = new List<SchoolYearSemester>();
                    SYSList.Add(SYS);
                    AutoSummaryList = AutoSummary.Select(allStudentID, SYSList);
                    AttendanceList = Attendance.SelectBySchoolYearAndSemester(K12.Data.Student.SelectByIDs(allStudentID), int.Parse(form.SchoolYear), int.Parse(form.Semester));
                }
            }

            //if (AttendanceList.Count == 0)
            //{
            //    MsgBox.Show("未取得缺曠資料");
            //    e.Cancel = true;
            //    return;
            //}

            foreach (AutoSummaryRecord var in AutoSummaryList)
            {
                if (!AutoSummaryDic.ContainsKey(var.RefStudentID))
                {
                    AutoSummaryDic.Add(var.RefStudentID, new List<AutoSummaryRecord>());
                }

                if (var.AbsenceCounts.Count > 0)
                {
                    AutoSummaryDic[var.RefStudentID].Add(var);
                }
            }

            foreach (List<AutoSummaryRecord> list in AutoSummaryDic.Values)
            {
                list.Sort(SortBySchoolYear);
            }

            AttendanceList.Sort(SortByOccurDate);

            foreach (AttendanceRecord var in AttendanceList)
            {
                #region AttendanceList
                string studentID = var.RefStudentID;
                string schoolYear = var.SchoolYear.ToString();
                string semester = var.Semester.ToString();
                string occurDate = var.OccurDate.ToShortDateString();
                string sso = schoolYear + "_" + semester + "_" + occurDate;

                //累計資料
                if (!periodStatisticsByType.ContainsKey(studentID))
                    periodStatisticsByType.Add(studentID, new Dictionary<string, Dictionary<string, int>>());
                //預設資料
                foreach (string value in periodList.Values)
                {
                    if (!periodStatisticsByType[studentID].ContainsKey(value))
                        periodStatisticsByType[studentID].Add(value, new Dictionary<string, int>());
                    foreach (string absence in absenceList.Keys)
                    {
                        if (!periodStatisticsByType[studentID][value].ContainsKey(absence))
                            periodStatisticsByType[studentID][value].Add(absence, 0);
                    }
                }

                //每一位學生缺曠資料
                if (!studentAbsenceDetail.ContainsKey(studentID))
                    studentAbsenceDetail.Add(studentID, new Dictionary<string, Dictionary<string, string>>());
                if (!studentAbsenceDetail[studentID].ContainsKey(sso))
                    studentAbsenceDetail[studentID].Add(sso, new Dictionary<string, string>());

                foreach (AttendancePeriod period in var.PeriodDetail)
                {
                    string absenceType = period.AbsenceType;
                    string inner = period.Period;
                    if (!periodList.ContainsKey(inner))
                        continue;
                    string periodType = periodList[inner];

                    if (!studentAbsenceDetail[studentID][sso].ContainsKey(inner) && absenceList.ContainsKey(absenceType))
                        studentAbsenceDetail[studentID][sso].Add(inner, absenceList[absenceType]);

                    if (periodStatisticsByType[studentID][periodType].ContainsKey(absenceType))
                        periodStatisticsByType[studentID][periodType][absenceType]++;
                }
                #endregion
            }

            #endregion

            #region 產生範本

            Workbook template = new Workbook(new MemoryStream(ProjectResource.學生缺曠明細));
            
            Range tempEachColumn = template.Worksheets[0].Cells.CreateRange(4, 1, true);

            Workbook prototype = new Workbook();
            prototype.Copy(template);
            prototype.CopyTheme(template);

            Worksheet ptws = prototype.Worksheets[0];
            int startPage = 1;
            int pageNumber = 1;

            int colIndex = 4;

            int startPeriodIndex = colIndex;
            int endPeriodIndex;

            //產生每一個節次的欄位
            foreach (string periodName in periodList.Keys)
            {
                tool.CopyStyle(ptws.Cells.CreateRange(colIndex, 1, true), tempEachColumn);
                ptws.Cells[4, colIndex].PutValue(periodName);
                columnTable.Add(periodName, colIndex);
                colIndex++;
            }
            endPeriodIndex = colIndex;

            ptws.Cells.CreateRange(3, startPeriodIndex, 1, endPeriodIndex - startPeriodIndex).Merge();
            ptws.Cells[3, startPeriodIndex].PutValue("節次");

            //合併標題列
            ptws.Cells.CreateRange(0, 0, 1, endPeriodIndex).Merge();
            ptws.Cells.CreateRange(1, 0, 1, endPeriodIndex).Merge();
            ptws.Cells.CreateRange(2, 0, 1, endPeriodIndex).Merge();

            Range ptHeader = ptws.Cells.CreateRange(0, 5, false);
            Range ptEachRow = ptws.Cells.CreateRange(5, 1, false);

            //current++;

            #endregion

            #region 產生報表

            Workbook wb = new Workbook();
            wb.Copy(prototype);
            wb.CopyTheme(prototype);

            Worksheet ws = wb.Worksheets[0];

            int index = 0;
            int dataIndex = 0;

            int studentCount = 1;

            foreach (JHStudentRecord studentInfo in selectedStudents)
            {
                #region selectedStudents
                string TitleName1 = School.ChineseName + "\n個人缺曠明細";
                string time_2013 = "";
                if (form.SelectDayOrSchoolYear)
                {
                    time_2013 = "統計區間：" + form.StartDay.ToShortDateString() + "~" + form.EndDay.ToShortDateString();
                }
                else
                {
                    if (form.checkBoxX1Bool) //全部學期列印
                    {
                        time_2013 = "統計區間：(所有學年度)";
                    }
                    else
                    {
                        time_2013 = string.Format("統計區間：{0}學年度 第{1}學期", form.SchoolYear, form.Semester);
                    }
                }
                string TitleName2 = "班級：" + ((studentInfo.Class == null ? "　" : studentInfo.Class.Name) + "　座號：" + ((studentInfo.SeatNo == null) ? "　" : studentInfo.SeatNo.ToString()) + "　學號：" + studentInfo.StudentNumber + "\n姓名：" + studentInfo.Name);
                string TitleName3 = time_2013;
                
                //回報進度
                _BGWAbsenceDetail.ReportProgress((int)(((double)studentCount++ * 100.0) / (double)selectedStudents.Count));

                //該名學生,同時有明細與非明細
                if (studentAbsenceDetail.ContainsKey(studentInfo.ID) && AutoSummaryDic.ContainsKey(studentInfo.ID))
                {
                    #region 明細列印

                    //如果不是第一頁，就在上一頁的資料列下邊加黑線
                    if (index != 0)
                        ws.Cells.CreateRange(index - 1, 0, 1, endPeriodIndex).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

                    //複製 Header
                    tool.CopyStyle(ws.Cells.CreateRange(index, 5, false), ptHeader);

                    dataIndex = index + 5;
                    int recordCount = 0;

                    //學生Row筆數超過40筆,則添加換行符號,與標頭
                    int CountRows = 0;

                    Dictionary<string, Dictionary<string, string>> absenceDetail = studentAbsenceDetail[studentInfo.ID];

                    //取總頁數 , 資料數除以38列(70/38=2)
                    int TotlePage = absenceDetail.Count / 40;
                    //目前頁數
                    int pageCount = 1;
                    //如果還有餘數則+1
                    if (absenceDetail.Count % 40 != 0)
                    {
                        TotlePage++;
                    }

                    //填寫基本資料
                    ws.Cells[index, 0].PutValue(TitleName1 + "(" + pageCount.ToString() + "/" + TotlePage.ToString() + ")");
                    pageCount++;
                    ws.Cells[index + 1, 0].PutValue(TitleName2);
                    ws.Cells[index + 2, 0].PutValue(TitleName3);

                    foreach (string sso in absenceDetail.Keys)
                    {
                        CountRows++;

                        string[] ssoSplit = sso.Split(new string[] { "_" }, StringSplitOptions.RemoveEmptyEntries);

                        //複製每一個 row
                        tool.CopyStyle(ws.Cells.CreateRange(dataIndex, 1, false), ptEachRow);

                        //填寫學生缺曠資料
                        ws.Cells[dataIndex, 0].PutValue(ssoSplit[0]);
                        ws.Cells[dataIndex, 1].PutValue(ssoSplit[1]);
                        ws.Cells[dataIndex, 2].PutValue(ssoSplit[2]);
                        ws.Cells[dataIndex, 3].PutValue(CommonMethods.GetChineseDayOfWeek(DateTime.Parse(ssoSplit[2])));

                        Dictionary<string, string> record = absenceDetail[sso];
                        foreach (string periodName in record.Keys)
                        {
                            ws.Cells[dataIndex, columnTable[periodName]].PutValue(record[periodName]);
                        }

                        dataIndex++;
                        recordCount++;

                        if (CountRows == 40 && pageCount <= TotlePage)
                        {
                            CountRows = 0;
                            //分頁
                            ws.HorizontalPageBreaks.Add(dataIndex, endPeriodIndex);
                            //複製 Header
                            tool.CopyStyle(ws.Cells.CreateRange(dataIndex, 5, false), ptHeader);
                            //填寫基本資料
                            ws.Cells[dataIndex, 0].PutValue(TitleName1 + "(" + pageCount.ToString() + "/" + TotlePage.ToString() + ")");
                            pageCount++; //下一頁使用
                            ws.Cells[dataIndex + 1, 0].PutValue(TitleName2);
                            ws.Cells[dataIndex + 2, 0].PutValue(TitleName3);
                            dataIndex += 5;
                        }

                    }

                    //缺曠統計資訊
                    Range absenceStatisticsRange = ws.Cells.CreateRange(dataIndex, 0, 1, endPeriodIndex);
                    absenceStatisticsRange.Merge();
                    //上
                    absenceStatisticsRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                    //左
                    absenceStatisticsRange.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                    //右
                    absenceStatisticsRange.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                    //Row的高度
                    absenceStatisticsRange.RowHeight = 14.0;
                    
                    tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                    tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                    tool.SetFontSize(ws.Cells[dataIndex, 0], 10);

                    tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                    tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                    tool.SetFontSize(ws.Cells[dataIndex, 0], 10);


                    //文字內容
                    ws.Cells[dataIndex, 0].PutValue("上列之缺曠明細加總為：");
                    dataIndex++;

                    foreach (string periodType in periodStatisticsByType[studentInfo.ID].Keys)
                    {
                        Dictionary<string, int> byType = periodStatisticsByType[studentInfo.ID][periodType];
                        int printable = 0;
                        foreach (string type in byType.Keys)
                        {
                            printable += byType[type];
                        }
                        if (printable == 0)
                            continue;

                        ws.Cells.CreateRange(dataIndex, 0, 1, 1).Merge();
                        ws.Cells.CreateRange(dataIndex, 1, 1, endPeriodIndex - 1).Merge();
                        
                        tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                        tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                        tool.SetFontSize(ws.Cells[dataIndex, 0], 10);

                        ws.Cells.CreateRange(dataIndex, 0, 1, endPeriodIndex).RowHeight = 14.0;
                        //左
                        ws.Cells.CreateRange(dataIndex, 0, 1, 1).SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                        ws.Cells[dataIndex, 0].PutValue(periodType);

                        StringBuilder text = new StringBuilder("");

                        foreach (string type in byType.Keys)
                        {
                            if (byType[type] > 0)
                            {
                                if (text.ToString() != "")
                                    text.Append("　");
                                text.Append(type + "：" + byType[type]);
                            }
                        }
                        
                        tool.SetFontSize(ws.Cells[dataIndex, 0], 10);
                        tool.SetShrinkToFit(ws.Cells[dataIndex, 0], true);

                        ws.Cells[dataIndex, 1].PutValue(text.ToString());
                        ws.Cells.CreateRange(dataIndex, 0, 1, endPeriodIndex).SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

                        dataIndex++;
                    }
                    #endregion

                    if (!form.SelectDayOrSchoolYear)
                    {
                        #region 如果是依學期進行列印
                        bool IsValue = false;
                        foreach (AutoSummaryRecord auto in AutoSummaryDic[studentInfo.ID])
                        {
                            if (auto.InitialSummary == null)
                                continue;

                            if (auto.InitialSummary.SelectSingleNode("AttendanceStatistics") == null)
                                continue;

                            int count = 0;
                            foreach (AbsenceCountRecord acr in auto.InitialAbsenceCounts)
                            {
                                count += acr.Count;
                            }
                            if (count > 0)
                            {
                                IsValue = true;
                            }
                        }

                        if (IsValue)
                        {
                            absenceStatisticsRange = ws.Cells.CreateRange(dataIndex, 0, 1, endPeriodIndex);
                            absenceStatisticsRange.Merge();
                            absenceStatisticsRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                            absenceStatisticsRange.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                            absenceStatisticsRange.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                            absenceStatisticsRange.RowHeight = 14.0;

                            tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                            tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                            tool.SetFontSize(ws.Cells[dataIndex, 0], 10);

                            ws.Cells[dataIndex, 0].PutValue("缺曠非明細(僅轉入生與特殊狀況之學生會以非明細記錄)：");

                            Dictionary<string, Dictionary<string, List<AbsenceCountRecord>>> AbsenceDic = new Dictionary<string, Dictionary<string, List<AbsenceCountRecord>>>();
                            #region 依學年度學期整理資料
                            foreach (AutoSummaryRecord auto in AutoSummaryDic[studentInfo.ID])
                            {
                                string sys = auto.SchoolYear.ToString() + "/" + auto.Semester.ToString();
                                foreach (AbsenceCountRecord acr in auto.InitialAbsenceCounts)
                                {
                                    if (!AbsenceDic.ContainsKey(sys))
                                    {
                                        AbsenceDic.Add(sys, new Dictionary<string, List<AbsenceCountRecord>>());
                                    }
                                    if (!AbsenceDic[sys].ContainsKey(acr.PeriodType))
                                    {
                                        AbsenceDic[sys].Add(acr.PeriodType, new List<AbsenceCountRecord>());
                                    }
                                    AbsenceDic[sys][acr.PeriodType].Add(acr);
                                }
                            }
                            #endregion
                            foreach (string list in AbsenceDic.Keys)
                            {
                                dataIndex++;

                                //學年度
                                absenceStatisticsRange = ws.Cells.CreateRange(dataIndex, 0, 1, endPeriodIndex);
                                absenceStatisticsRange.Merge();
                                absenceStatisticsRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.None, Color.Transparent);
                                absenceStatisticsRange.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.None, Color.Transparent);
                                absenceStatisticsRange.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                                absenceStatisticsRange.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                                absenceStatisticsRange.RowHeight = 14.0;
                                
                                tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                                tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                                tool.SetFontSize(ws.Cells[dataIndex, 0], 10);

                                ws.Cells[dataIndex, 0].PutValue(list);

                                //一般
                                foreach (string list1 in AbsenceDic[list].Keys)
                                {
                                    dataIndex++;
                                    absenceStatisticsRange = ws.Cells.CreateRange(dataIndex, 0, 1, 1);
                                    ws.Cells.CreateRange(dataIndex, 1, 1, endPeriodIndex - 1).Merge();
                                    absenceStatisticsRange.Merge();
                                    absenceStatisticsRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.None, Color.Transparent);
                                    absenceStatisticsRange.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.None, Color.Transparent);
                                    absenceStatisticsRange.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                                    absenceStatisticsRange.RowHeight = 14.0;

                                    tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                                    tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                                    tool.SetFontSize(ws.Cells[dataIndex, 0], 10);

                                    ws.Cells[dataIndex, 0].PutValue(list1); //一般

                                    StringBuilder text = new StringBuilder();
                                    foreach (AbsenceCountRecord acrr in AbsenceDic[list][list1])
                                    {
                                        text.Append(acrr.Name + "：" + acrr.Count.ToString() + "　");
                                    }

                                    absenceStatisticsRange = ws.Cells.CreateRange(dataIndex, 0, 1, endPeriodIndex);
                                    absenceStatisticsRange.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                                    
                                    tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                                    tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                                    tool.SetFontSize(ws.Cells[dataIndex, 0], 10);

                                    ws.Cells[dataIndex, 1].PutValue(text.ToString()); //統計
                                }
                            }
                        }
                        else
                        {
                            dataIndex--;
                        }
                        dataIndex++;
                        #endregion
                    }

                } //有明細但無非明細
                else if (studentAbsenceDetail.ContainsKey(studentInfo.ID) && !AutoSummaryDic.ContainsKey(studentInfo.ID))
                {
                    #region 明細列印

                    //如果不是第一頁，就在上一頁的資料列下邊加黑線
                    if (index != 0)
                        ws.Cells.CreateRange(index - 1, 0, 1, endPeriodIndex).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

                    //複製 Header
                    tool.CopyStyle(ws.Cells.CreateRange(index, 5, false), ptHeader);

                    dataIndex = index + 5;
                    int recordCount = 0;

                    //學生Row筆數超過40筆,則添加換行符號,與標頭
                    int CountRows = 0;

                    Dictionary<string, Dictionary<string, string>> absenceDetail = studentAbsenceDetail[studentInfo.ID];

                    //取總頁數 , 資料數除以38列(70/38=2)
                    int TotlePage = absenceDetail.Count / 40;
                    //目前頁數
                    int pageCount = 1;
                    //如果還有餘數則+1
                    if (absenceDetail.Count % 40 != 0)
                    {
                        TotlePage++;
                    }

                    //填寫基本資料
                    ws.Cells[index, 0].PutValue(TitleName1 + "(" + pageCount.ToString() + "/" + TotlePage.ToString() + ")");
                    pageCount++;
                    ws.Cells[index + 1, 0].PutValue(TitleName2);
                    ws.Cells[index + 2, 0].PutValue(TitleName3);

                    foreach (string sso in absenceDetail.Keys)
                    {
                        CountRows++;

                        string[] ssoSplit = sso.Split(new string[] { "_" }, StringSplitOptions.RemoveEmptyEntries);

                        //複製每一個 row
                        tool.CopyStyle(ws.Cells.CreateRange(dataIndex, 1, false), ptEachRow);

                        //填寫學生缺曠資料
                        ws.Cells[dataIndex, 0].PutValue(ssoSplit[0]);
                        ws.Cells[dataIndex, 1].PutValue(ssoSplit[1]);
                        ws.Cells[dataIndex, 2].PutValue(ssoSplit[2]);
                        ws.Cells[dataIndex, 3].PutValue(CommonMethods.GetChineseDayOfWeek(DateTime.Parse(ssoSplit[2])));

                        Dictionary<string, string> record = absenceDetail[sso];
                        foreach (string periodName in record.Keys)
                        {
                            ws.Cells[dataIndex, columnTable[periodName]].PutValue(record[periodName]);
                        }

                        dataIndex++;
                        recordCount++;

                        if (CountRows == 40 && pageCount <= TotlePage)
                        {
                            CountRows = 0;
                            //分頁
                            ws.HorizontalPageBreaks.Add(dataIndex, endPeriodIndex);
                            //複製 Header
                            tool.CopyStyle(ws.Cells.CreateRange(dataIndex, 5, false), ptHeader);

                            //填寫基本資料
                            ws.Cells[dataIndex, 0].PutValue(TitleName1 + "(" + pageCount.ToString() + "/" + TotlePage.ToString() + ")");
                            pageCount++; //下一頁使用
                            ws.Cells[dataIndex + 1, 0].PutValue(TitleName2);
                            ws.Cells[dataIndex + 2, 0].PutValue(TitleName3);
                            dataIndex += 5;
                        }

                    }
                    #endregion

                    #region 插入獎懲統計資訊
                    Range absenceStatisticsRange = ws.Cells.CreateRange(dataIndex, 0, 1, endPeriodIndex);
                    absenceStatisticsRange.Merge();
                    //上
                    absenceStatisticsRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                    //左
                    absenceStatisticsRange.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                    //右
                    absenceStatisticsRange.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                    //Row的高度
                    absenceStatisticsRange.RowHeight = 14.0;

                    tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                    tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                    tool.SetFontSize(ws.Cells[dataIndex, 0], 10);
                    
                    //文字內容
                    ws.Cells[dataIndex, 0].PutValue("上列之缺曠明細加總為：");
                    dataIndex++;

                    foreach (string periodType in periodStatisticsByType[studentInfo.ID].Keys)
                    {
                        Dictionary<string, int> byType = periodStatisticsByType[studentInfo.ID][periodType];
                        int printable = 0;
                        foreach (string type in byType.Keys)
                        {
                            printable += byType[type];
                        }
                        if (printable == 0)
                            continue;

                        ws.Cells.CreateRange(dataIndex, 0, 1, 1).Merge();
                        ws.Cells.CreateRange(dataIndex, 1, 1, endPeriodIndex - 1).Merge();

                        tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                        tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                        tool.SetFontSize(ws.Cells[dataIndex, 0], 10);

                        ws.Cells.CreateRange(dataIndex, 0, 1, endPeriodIndex).RowHeight = 14.0;
                        //左
                        ws.Cells.CreateRange(dataIndex, 0, 1, 1).SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                        ws.Cells[dataIndex, 0].PutValue(periodType);

                        StringBuilder text = new StringBuilder("");

                        foreach (string type in byType.Keys)
                        {
                            if (byType[type] > 0)
                            {
                                if (text.ToString() != "")
                                    text.Append("　");
                                text.Append(type + "：" + byType[type]);
                            }
                        }

                        tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                        tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);

                        tool.SetFontSize(ws.Cells[dataIndex, 0], 10);
                        tool.SetShrinkToFit(ws.Cells[dataIndex, 0], true);

                        ws.Cells[dataIndex, 1].PutValue(text.ToString());
                        ws.Cells.CreateRange(dataIndex, 0, 1, endPeriodIndex).SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

                        dataIndex++;
                    }
                    #endregion

                } //無明細但有非明細
                else if (!studentAbsenceDetail.ContainsKey(studentInfo.ID) && AutoSummaryDic.ContainsKey(studentInfo.ID))
                {
                    if (AutoSummaryDic[studentInfo.ID].Count == 0)
                        continue;

                    if (!form.SelectDayOrSchoolYear)
                    {
                        #region 如果是依學期進行列印

                        bool IsValue = false;
                        foreach (AutoSummaryRecord auto in AutoSummaryDic[studentInfo.ID])
                        {
                            int count = 0;
                            foreach (AbsenceCountRecord acr in auto.InitialAbsenceCounts)
                            {
                                count += acr.Count;
                            }
                            if (count > 0)
                            {
                                IsValue = true;
                            }
                        }

                        if (IsValue)
                        {
                            //如果不是第一頁，就在上一頁的資料列下邊加黑線
                            if (index != 0)
                                ws.Cells.CreateRange(index - 1, 0, 1, endPeriodIndex).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

                            //複製 Header
                            tool.CopyStyle(ws.Cells.CreateRange(index, 5, false), ptHeader);

                            dataIndex = index + 5;

                            //填寫基本資料
                            ws.Cells[index, 0].PutValue(TitleName1 + "(1/1)");
                            ws.Cells[index + 1, 0].PutValue(TitleName2);
                            ws.Cells[index + 2, 0].PutValue(TitleName3);
                            #region 缺曠統計資訊
                            Range absenceStatisticsRange = ws.Cells.CreateRange(dataIndex, 0, 1, endPeriodIndex);
                            absenceStatisticsRange.Merge();
                            //上
                            absenceStatisticsRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                            //左
                            absenceStatisticsRange.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                            //右
                            absenceStatisticsRange.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                            //Row的高度
                            absenceStatisticsRange.RowHeight = 14.0;
                            
                            tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                            tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                            tool.SetFontSize(ws.Cells[dataIndex, 0], 10);

                            //文字內容
                            ws.Cells[dataIndex, 0].PutValue("(無明細資料)");
                            dataIndex++;
                            #endregion

                            absenceStatisticsRange = ws.Cells.CreateRange(dataIndex, 0, 1, endPeriodIndex);
                            absenceStatisticsRange.Merge();
                            absenceStatisticsRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                            absenceStatisticsRange.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                            absenceStatisticsRange.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                            absenceStatisticsRange.RowHeight = 14.0;

                            tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                            tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                            tool.SetFontSize(ws.Cells[dataIndex, 0], 10);

                            ws.Cells[dataIndex, 0].PutValue("缺曠非明細(僅轉入生與特殊狀況之學生會以非明細記錄)：");

                            Dictionary<string, Dictionary<string, List<AbsenceCountRecord>>> AbsenceDic = new Dictionary<string, Dictionary<string, List<AbsenceCountRecord>>>();
                            #region 依學年度學期整理資料
                            foreach (AutoSummaryRecord auto in AutoSummaryDic[studentInfo.ID])
                            {
                                string sys = auto.SchoolYear.ToString() + "/" + auto.Semester.ToString();
                                foreach (AbsenceCountRecord acr in auto.InitialAbsenceCounts)
                                {
                                    if (!AbsenceDic.ContainsKey(sys))
                                    {
                                        AbsenceDic.Add(sys, new Dictionary<string, List<AbsenceCountRecord>>());
                                    }
                                    if (!AbsenceDic[sys].ContainsKey(acr.PeriodType))
                                    {
                                        AbsenceDic[sys].Add(acr.PeriodType, new List<AbsenceCountRecord>());
                                    }
                                    AbsenceDic[sys][acr.PeriodType].Add(acr);
                                }
                            }
                            #endregion
                            foreach (string list in AbsenceDic.Keys)
                            {
                                dataIndex++;

                                //學年度
                                absenceStatisticsRange = ws.Cells.CreateRange(dataIndex, 0, 1, endPeriodIndex);
                                absenceStatisticsRange.Merge();
                                absenceStatisticsRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.None, Color.Transparent);
                                absenceStatisticsRange.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.None, Color.Transparent);
                                absenceStatisticsRange.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                                absenceStatisticsRange.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                                absenceStatisticsRange.RowHeight = 14.0;
                                
                                tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                                tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                                tool.SetFontSize(ws.Cells[dataIndex, 0], 10);

                                ws.Cells[dataIndex, 0].PutValue(list);

                                //一般
                                foreach (string list1 in AbsenceDic[list].Keys)
                                {
                                    dataIndex++;
                                    absenceStatisticsRange = ws.Cells.CreateRange(dataIndex, 0, 1, 1);
                                    ws.Cells.CreateRange(dataIndex, 1, 1, endPeriodIndex - 1).Merge();
                                    absenceStatisticsRange.Merge();
                                    absenceStatisticsRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.None, Color.Transparent);
                                    absenceStatisticsRange.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.None, Color.Transparent);
                                    absenceStatisticsRange.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                                    absenceStatisticsRange.RowHeight = 14.0;

                                    tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                                    tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                                    tool.SetFontSize(ws.Cells[dataIndex, 0], 10);

                                    ws.Cells[dataIndex, 0].PutValue(list1); //一般

                                    StringBuilder text = new StringBuilder();
                                    foreach (AbsenceCountRecord acrr in AbsenceDic[list][list1])
                                    {
                                        text.Append(acrr.Name + "：" + acrr.Count.ToString() + "　");
                                    }

                                    absenceStatisticsRange = ws.Cells.CreateRange(dataIndex, 0, 1, endPeriodIndex);
                                    absenceStatisticsRange.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                                    
                                    tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                                    tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                                    tool.SetFontSize(ws.Cells[dataIndex, 0], 10);

                                    ws.Cells[dataIndex, 1].PutValue(text.ToString()); //統計
                                }
                            }
                        }
                        else
                        {
                            dataIndex--;
                        }
                        dataIndex++;
                        #endregion
                    }
                }
                else
                { //皆無資料
                    continue;
                }

                index = dataIndex;

                //每500頁,增加一個WorkSheet,並於下標顯示(1~500)(501~xxx)
                if (pageNumber < 500)
                {
                    try
                    {
                        ws.HorizontalPageBreaks.Add(index, endPeriodIndex);
                        pageNumber++;
                    }
                    catch
                    {
                    }
                }
                else
                {
                    ws.Name = startPage + " ~ " + (pageNumber + startPage - 1);
                    ws = wb.Worksheets[wb.Worksheets.Add()];
                    ws.Copy(prototype.Worksheets[0]);

                    startPage += pageNumber;
                    pageNumber = 1;
                    index = 0;
                }
                #endregion
            }

            if (dataIndex > 0)
            {
                //最後一頁的資料列下邊加上黑線
                ws.Cells.CreateRange(dataIndex - 1, 0, 1, endPeriodIndex).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                ws.Name = startPage + " ~ " + (pageNumber + startPage - 2);
            }
            else
                wb = new Workbook();

            #endregion

            string path = Path.Combine(Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".xlsx");
            e.Result = new object[] { reportName, path, wb };
        }

        public int SortBySchoolYear(AutoSummaryRecord ASR1, AutoSummaryRecord ASR2)
        {
            string ASR1String = ASR1.SchoolYear.ToString().PadLeft(3, '0') + ASR1.Semester.ToString();
            string ASR2String = ASR2.SchoolYear.ToString().PadLeft(3, '0') + ASR2.Semester.ToString();
            return ASR1String.CompareTo(ASR2String);
        }

        public int SortByOccurDate(AttendanceRecord ASR1, AttendanceRecord ASR2)
        {
            return ASR1.OccurDate.CompareTo(ASR2.OccurDate);
        }
    }
}
