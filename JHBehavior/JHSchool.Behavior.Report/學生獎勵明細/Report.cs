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

namespace JHSchool.Behavior.Report.學生獎勵明細
{
    internal class Report : IReport
    {
        #region IReport 成員

        private BackgroundWorker _BGWDisciplineDetail;
        SelectMeritForm form;

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

            form = new SelectMeritForm();

            if (form.ShowDialog() == DialogResult.OK)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化學生獎勵明細...");

                //object[] args = new object[] { form.SchoolYear, form.Semester };

                _BGWDisciplineDetail = new BackgroundWorker();
                _BGWDisciplineDetail.DoWork += new DoWorkEventHandler(_BGWDisciplineDetail_DoWork);
                _BGWDisciplineDetail.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CommonMethods.ExcelReport_RunWorkerCompleted);
                _BGWDisciplineDetail.ProgressChanged += new ProgressChangedEventHandler(CommonMethods.Report_ProgressChanged);
                _BGWDisciplineDetail.WorkerReportsProgress = true;
                _BGWDisciplineDetail.RunWorkerAsync();
            }
        }

        #endregion

        void _BGWDisciplineDetail_DoWork(object sender, DoWorkEventArgs e)
        {
            string reportName = "學生獎勵明細";

            #region 快取相關資料

            //選擇的學生
            List<JHStudentRecord> selectedStudents = JHStudent.SelectByIDs(Student.Instance.SelectedKeys);
            selectedStudents = SortClassIndex.JHSchoolData_JHStudentRecord(selectedStudents);

            //紀錄所有學生ID
            List<string> allStudentID = new List<string>();

            //每一位學生的獎勵明細
            Dictionary<string, Dictionary<string, Dictionary<string, string>>> studentDisciplineDetail = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();

            //每一位學生的獎勵累計資料
            Dictionary<string, Dictionary<string, int>> studentDisciplineStatistics = new Dictionary<string, Dictionary<string, int>>();

            //每一位學生ID,會有多筆的AutoSummary
            Dictionary<string, List<AutoSummaryRecord>> AutoSummaryDic = new Dictionary<string, List<AutoSummaryRecord>>();

            //紀錄每一種獎勵在報表中的 column index
            Dictionary<string, int> columnTable = new Dictionary<string, int>();

            //取得所有學生ID
            foreach (JHStudentRecord var in selectedStudents)
            {
                allStudentID.Add(var.ID);
            }

            //對照表
            Dictionary<string, string> meritTable = new Dictionary<string, string>();
            meritTable.Add("A", "大功");
            meritTable.Add("B", "小功");
            meritTable.Add("C", "嘉獎");

            //初始化
            string[] columnString = new string[] { "嘉獎", "小功", "大功", "事由", "備註" };
            int i = 4;
            foreach (string s in columnString)
            {
                columnTable.Add(s, i++);
            }

            //取得獎懲資料
            List<MeritRecord> MeritList = new List<MeritRecord>();

            //取得自動統計資料
            List<AutoSummaryRecord> AutoSummaryList = new List<AutoSummaryRecord>();

            if (form.SelectDayOrSchoolYear) //依日期
            {
                if (form.SetupTime) //依發生日期
                {
                    MeritList = Merit.SelectByOccurDate(allStudentID, form.StartDateTime, form.EndDateTime);
                }
                else //依登錄日期
                {
                    MeritList = Merit.SelectByRegisterDate(allStudentID, form.StartDateTime, form.EndDateTime);
                }
            }
            else //依學期
            {
                if (form.checkBoxX1Bool) //全部學期列印
                {
                    AutoSummaryList = AutoSummary.Select(allStudentID, null);
                    MeritList = Merit.SelectByStudentIDs(allStudentID);
                }
                else //指定學期列印
                {
                    SchoolYearSemester SYS = new SchoolYearSemester(int.Parse(form.SchoolYear), int.Parse(form.Semester));
                    List<SchoolYearSemester> SYSList = new List<SchoolYearSemester>();
                    SYSList.Add(SYS);
                    AutoSummaryList = AutoSummary.Select(allStudentID, SYSList);
                    foreach (MeritRecord each in Merit.SelectByStudentIDs(allStudentID))
                    {
                        //因為Discipline未提供依學期取得資料,只好自己判斷
                        if (each.SchoolYear.ToString() == form.SchoolYear && each.Semester.ToString() == form.Semester)
                        {
                            MeritList.Add(each);
                        }
                    }
                }
            }

            //if (MeritList.Count == 0)
            //{
            //    MsgBox.Show("未取得獎勵資料");
            //    e.Cancel = true;
            //    return;
            //}

            foreach (AutoSummaryRecord var in AutoSummaryList)
            {
                if (!AutoSummaryDic.ContainsKey(var.RefStudentID))
                {
                    AutoSummaryDic.Add(var.RefStudentID, new List<AutoSummaryRecord>());
                }

                //如果該學期獎勵與懲戒皆有數值
                if (var.MeritA + var.MeritB + var.MeritC > 0)
                {
                    AutoSummaryDic[var.RefStudentID].Add(var);
                }
            }

            foreach (List<AutoSummaryRecord> list in AutoSummaryDic.Values)
            {
                list.Sort(SortBySchoolYear);
            }

            MeritList.Sort(SortByOccurDate);

            foreach (MeritRecord var in MeritList)
            {
                #region MeritList
                string studentID = var.RefStudentID;
                string schoolYear = var.SchoolYear.ToString();
                string semester = var.Semester.ToString();
                string occurDate = var.OccurDate.ToShortDateString();
                string reason = var.Reason;
                string remark = var.Remark;
                string disciplineID = var.ID;
                string sso = schoolYear + "_" + semester + "_" + occurDate + "_" + disciplineID;

                //初始化累計資料
                if (!studentDisciplineStatistics.ContainsKey(studentID))
                {
                    studentDisciplineStatistics.Add(studentID, new Dictionary<string, int>());
                    studentDisciplineStatistics[studentID].Add("大功", 0);
                    studentDisciplineStatistics[studentID].Add("小功", 0);
                    studentDisciplineStatistics[studentID].Add("嘉獎", 0);
                }

                //每一位學生獎勵資料
                if (!studentDisciplineDetail.ContainsKey(studentID))
                    studentDisciplineDetail.Add(studentID, new Dictionary<string, Dictionary<string, string>>());
                if (!studentDisciplineDetail[studentID].ContainsKey(sso))
                    studentDisciplineDetail[studentID].Add(sso, new Dictionary<string, string>());

                //加入事由
                if (!studentDisciplineDetail[studentID][sso].ContainsKey("事由"))
                    studentDisciplineDetail[studentID][sso].Add("事由", reason);

                //加入備註
                if (!studentDisciplineDetail[studentID][sso].ContainsKey("備註"))
                    studentDisciplineDetail[studentID][sso].Add("備註", remark);

                studentDisciplineStatistics[studentID]["大功"] += var.MeritA.Value;
                studentDisciplineStatistics[studentID]["小功"] += var.MeritB.Value;
                studentDisciplineStatistics[studentID]["嘉獎"] += var.MeritC.Value;

                //明細資料
                studentDisciplineDetail[studentID][sso].Add("大功", var.MeritA.Value.ToString());
                studentDisciplineDetail[studentID][sso].Add("小功", var.MeritB.Value.ToString());
                studentDisciplineDetail[studentID][sso].Add("嘉獎", var.MeritC.Value.ToString());

                #endregion
            }

            #endregion

            #region 產生範本

            Workbook template = new Workbook(new MemoryStream(ProjectResource.學生獎勵記錄明細));

            Workbook prototype = new Workbook();
            prototype.Copy(template);
            prototype.CopyTheme(template);

            Worksheet ptws = prototype.Worksheets[0];

            int startPage = 1;
            int pageNumber = 1;

            int columnNumber = 9;

            //合併標題列
            ptws.Cells.CreateRange(0, 0, 1, columnNumber).Merge();
            ptws.Cells.CreateRange(1, 0, 1, columnNumber).Merge();

            Range ptHeader = ptws.Cells.CreateRange(0, 4, false);
            Range ptEachRow = ptws.Cells.CreateRange(4, 1, false);

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
                string TitleName1 = School.ChineseName + "\n個人獎勵明細";
                string TitleName2 = "班級：" + ((studentInfo.Class == null ? "　　　" : studentInfo.Class.Name) + "　　座號：" + ((studentInfo.SeatNo == null) ? "　" : studentInfo.SeatNo.ToString()) + "　　姓名：" + studentInfo.Name + "　　學號：" + studentInfo.StudentNumber);

                //回報進度
                _BGWDisciplineDetail.ReportProgress((int)(((double)studentCount++ * 100.0) / (double)selectedStudents.Count));

                //該名學生,同時有明細與非明細
                if (studentDisciplineStatistics.ContainsKey(studentInfo.ID) && AutoSummaryDic.ContainsKey(studentInfo.ID))
                {
                    #region 明細
                    //如果不是第一頁，就在上一頁的資料列下邊加黑線
                    if (index != 0)
                        ws.Cells.CreateRange(index - 1, 0, 1, columnNumber).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

                    //複製 Header
                    tool.CopyStyle(ws.Cells.CreateRange(index, 4, false), ptHeader);

                    dataIndex = index + 4;
                    int recordCount = 0;

                    //學生Row筆數超過40筆,則添加換行符號,與標頭
                    int CountRows = 0;

                    Dictionary<string, Dictionary<string, string>> disciplineDetail = studentDisciplineDetail[studentInfo.ID];

                    //取總頁數 , 資料數除以38列(70/38=2)
                    int TotlePage = disciplineDetail.Count / 40;
                    //目前頁數
                    int pageCount = 1;
                    //如果還有餘數則+1
                    if (disciplineDetail.Count % 40 != 0)
                    {
                        TotlePage++;
                    }

                    //填寫基本資料
                    ws.Cells[index, 0].PutValue(TitleName1 + "(" + pageCount.ToString() + "/" + TotlePage.ToString() + ")");
                    pageCount++;
                    ws.Cells[index + 1, 0].PutValue(TitleName2);

                    foreach (string sso in disciplineDetail.Keys)
                    {
                        CountRows++;

                        string[] ssoSplit = sso.Split(new string[] { "_" }, StringSplitOptions.RemoveEmptyEntries);

                        //複製每一個 row
                        tool.CopyStyle(ws.Cells.CreateRange(dataIndex, 1, false), ptEachRow);
                        //填寫學生獎勵資料
                        ws.Cells[dataIndex, 0].PutValue(ssoSplit[0]);
                        ws.Cells[dataIndex, 1].PutValue(ssoSplit[1]);
                        ws.Cells[dataIndex, 2].PutValue(ssoSplit[2]);
                        ws.Cells[dataIndex, 3].PutValue(CommonMethods.GetChineseDayOfWeek(DateTime.Parse(ssoSplit[2])));

                        Dictionary<string, string> record = disciplineDetail[sso];
                        foreach (string name in record.Keys)
                        {
                            if (meritTable.ContainsValue(name))
                            {
                                int v;

                                if (int.TryParse(record[name], out v))
                                {
                                    if (v > 0)
                                        ws.Cells[dataIndex, columnTable[name]].PutValue(record[name]);
                                }
                            }
                            else
                            {
                                if (columnTable.ContainsKey(name))
                                    ws.Cells[dataIndex, columnTable[name]].PutValue(record[name]);
                            }
                        }

                        dataIndex++;
                        recordCount++;

                        if (CountRows == 40 && pageCount <= TotlePage)
                        {
                            CountRows = 0;
                            //分頁
                            ws.HorizontalPageBreaks.Add(dataIndex, columnNumber);
                            //複製 Header
                            tool.CopyStyle(ws.Cells.CreateRange(dataIndex, 4, false), ptHeader);

                            //填寫基本資料
                            ws.Cells[dataIndex, 0].PutValue(TitleName1 + "(" + pageCount.ToString() + "/" + TotlePage.ToString() + ")");
                            pageCount++; //下一頁使用
                            ws.Cells[dataIndex + 1, 0].PutValue(TitleName2);

                            dataIndex += 4;
                        }
                    }

                    #endregion

                    #region 插入獎懲統計資訊
                    //獎懲統計資訊
                    Range disciplineStatisticsRange = ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber);
                    tool.CopyStyle(disciplineStatisticsRange, ptEachRow);

                    disciplineStatisticsRange.Merge();
                    //disciplineStatisticsRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Double, Color.Black);
                    //disciplineStatisticsRange.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Double, Color.Black);
                    disciplineStatisticsRange.RowHeight = 14.0;
                    
                    tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                    tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                    tool.SetFontSize(ws.Cells[dataIndex, 0], 10);


                    StringBuilder text = new StringBuilder("");
                    Dictionary<string, int> disciplineStatistics = studentDisciplineStatistics[studentInfo.ID];

                    foreach (string type in disciplineStatistics.Keys)
                    {
                        if (disciplineStatistics[type] > 0)
                        {
                            if (text.ToString() != "")
                                text.Append("　");
                            text.Append(type + "：" + disciplineStatistics[type]);
                        }
                    }
                    ws.Cells[dataIndex, 0].PutValue("上列之獎勵明細加總為：" + text.ToString());
                    dataIndex++;

                    #endregion 

                    if (AutoSummaryDic[studentInfo.ID].Count == 0)
                        continue;

                    //如果是依學期列印
                    if (!form.SelectDayOrSchoolYear)
                    {
                        #region 插入非明細資訊

                        bool IsValue = false;
                        foreach (AutoSummaryRecord auto in AutoSummaryDic[studentInfo.ID])
                        {
                            if (auto.InitialSummary == null)
                                continue;

                            if (auto.InitialSummary.SelectSingleNode("DisciplineStatistics") == null)
                                continue;

                            if (auto.InitialSummary.SelectSingleNode("DisciplineStatistics/Merit") != null)
                            {
                                if (auto.InitialMeritA + auto.InitialMeritB + auto.InitialMeritC > 0)
                                {
                                    IsValue = true;
                                }
                            }
                        }

                        if (IsValue)
                        {

                            //獎懲統計資訊
                            disciplineStatisticsRange = ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber);
                            tool.CopyStyle(disciplineStatisticsRange, ptEachRow);

                            disciplineStatisticsRange.Merge();
                            //disciplineStatisticsRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Double, Color.Black);
                            disciplineStatisticsRange.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.None, Color.Transparent);
                            //disciplineStatisticsRange.RowHeight = 14.0;
                            
                            tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                            tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                            tool.SetFontSize(ws.Cells[dataIndex, 0], 10);

                            ws.Cells[dataIndex, 0].PutValue("獎勵非明細(僅轉入生與特殊狀況之學生會以非明細記錄)：");

                            foreach (AutoSummaryRecord auto in AutoSummaryDic[studentInfo.ID])
                            {
                                if (auto.InitialSummary == null)
                                    continue;

                                if (auto.InitialSummary.SelectSingleNode("DisciplineStatistics") == null)
                                    continue;

                                dataIndex++;
                                text = new StringBuilder("");
                                text.Append(auto.SchoolYear.ToString() + "/" + auto.Semester.ToString() + "　");

                                if (auto.InitialSummary.SelectSingleNode("DisciplineStatistics/Merit") != null)
                                {
                                    if (auto.InitialMeritA + auto.InitialMeritB + auto.InitialMeritC > 0)
                                    {
                                        if (auto.InitialMeritA > 0)
                                            text.Append("大功：" + auto.InitialMeritA + "　");
                                        if (auto.InitialMeritB > 0)
                                            text.Append("小功：" + auto.InitialMeritB + "　");
                                        if (auto.InitialMeritC > 0)
                                            text.Append("嘉獎：" + auto.InitialMeritC + "　");
                                    }
                                }

                                //獎懲統計內容
                                tool.CopyStyle(ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber), ptEachRow);

                                ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber).RowHeight = 14.0;
                                ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber).Merge();
                                ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber).SetOutlineBorder(BorderType.TopBorder, CellBorderType.None, Color.Transparent);
                                ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.None, Color.Transparent);
                                
                                tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                                tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                                tool.SetFontSize(ws.Cells[dataIndex, 0], 10);
                                tool.SetShrinkToFit(ws.Cells[dataIndex, 0], true);

                                ws.Cells[dataIndex, 0].PutValue(text.ToString());
                            }

                            ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                            ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber).SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                            dataIndex++;
                        }
                        #endregion
                    }
                }
                else if (studentDisciplineStatistics.ContainsKey(studentInfo.ID) && !AutoSummaryDic.ContainsKey(studentInfo.ID))
                {
                    #region 明細
                    //如果不是第一頁，就在上一頁的資料列下邊加黑線
                    if (index != 0)
                        ws.Cells.CreateRange(index - 1, 0, 1, columnNumber).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

                    //複製 Header
                    tool.CopyStyle(ws.Cells.CreateRange(index, 4, false), ptHeader);

                    dataIndex = index + 4;
                    int recordCount = 0;

                    //學生Row筆數超過40筆,則添加換行符號,與標頭
                    int CountRows = 0;

                    Dictionary<string, Dictionary<string, string>> disciplineDetail = studentDisciplineDetail[studentInfo.ID];

                    //取總頁數 , 資料數除以38列(70/38=2)
                    int TotlePage = disciplineDetail.Count / 40;
                    //目前頁數
                    int pageCount = 1;
                    //如果還有餘數則+1
                    if (disciplineDetail.Count % 40 != 0)
                    {
                        TotlePage++;
                    }

                    //填寫基本資料
                    ws.Cells[index, 0].PutValue(TitleName1 + "(" + pageCount.ToString() + "/" + TotlePage.ToString() + ")");
                    pageCount++;
                    ws.Cells[index + 1, 0].PutValue(TitleName2);

                    foreach (string sso in disciplineDetail.Keys)
                    {
                        CountRows++;

                        string[] ssoSplit = sso.Split(new string[] { "_" }, StringSplitOptions.RemoveEmptyEntries);

                        //複製每一個 row
                        tool.CopyStyle(ws.Cells.CreateRange(dataIndex, 1, false), ptEachRow);

                        //填寫學生獎勵資料
                        ws.Cells[dataIndex, 0].PutValue(ssoSplit[0]);
                        ws.Cells[dataIndex, 1].PutValue(ssoSplit[1]);
                        ws.Cells[dataIndex, 2].PutValue(ssoSplit[2]);
                        ws.Cells[dataIndex, 3].PutValue(CommonMethods.GetChineseDayOfWeek(DateTime.Parse(ssoSplit[2])));

                        Dictionary<string, string> record = disciplineDetail[sso];
                        foreach (string name in record.Keys)
                        {
                            if (meritTable.ContainsValue(name))
                            {
                                int v;

                                if (int.TryParse(record[name], out v))
                                {
                                    if (v > 0)
                                        ws.Cells[dataIndex, columnTable[name]].PutValue(record[name]);
                                }
                            }
                            else
                            {
                                if (columnTable.ContainsKey(name))
                                    ws.Cells[dataIndex, columnTable[name]].PutValue(record[name]);
                            }
                        }

                        dataIndex++;
                        recordCount++;

                        if (CountRows == 40 && pageCount <= TotlePage)
                        {
                            CountRows = 0;
                            //分頁
                            ws.HorizontalPageBreaks.Add(dataIndex, columnNumber);
                            //複製 Header
                            tool.CopyStyle(ws.Cells.CreateRange(dataIndex, 4, false), ptHeader);

                            //填寫基本資料
                            ws.Cells[dataIndex, 0].PutValue(TitleName1 + "(" + pageCount.ToString() + "/" + TotlePage.ToString() + ")");
                            pageCount++; //下一頁使用
                            ws.Cells[dataIndex + 1, 0].PutValue(TitleName2);

                            dataIndex += 4;
                        }
                    }
                    #endregion

                    #region 插入獎懲統計資訊
                    //獎懲統計資訊
                    Range disciplineStatisticsRange = ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber);
                    tool.CopyStyle(disciplineStatisticsRange, ptEachRow);

                    disciplineStatisticsRange.Merge();
                    //disciplineStatisticsRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Double, Color.Black);
                    //disciplineStatisticsRange.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Double, Color.Black);
                    disciplineStatisticsRange.RowHeight = 14.0;
                    
                    tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                    tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                    tool.SetFontSize(ws.Cells[dataIndex, 0], 10);

                    StringBuilder text = new StringBuilder("");
                    Dictionary<string, int> disciplineStatistics = studentDisciplineStatistics[studentInfo.ID];

                    foreach (string type in disciplineStatistics.Keys)
                    {
                        if (disciplineStatistics[type] > 0)
                        {
                            if (text.ToString() != "")
                                text.Append("　");
                            text.Append(type + "：" + disciplineStatistics[type]);
                        }
                    }
                    ws.Cells[dataIndex, 0].PutValue("上列之獎勵明細加總為：" + text.ToString());
                    dataIndex++;

                    #endregion
                }
                else if (!studentDisciplineStatistics.ContainsKey(studentInfo.ID) && AutoSummaryDic.ContainsKey(studentInfo.ID))
                {
                    if (AutoSummaryDic[studentInfo.ID].Count == 0)
                        continue;

                    if (!form.SelectDayOrSchoolYear)
                    {
                        #region 如果是依學期進行列印

                        bool IsValue = false;
                        foreach (AutoSummaryRecord auto in AutoSummaryDic[studentInfo.ID])
                        {
                            if (auto.InitialSummary == null)
                                continue;

                            if (auto.InitialSummary.SelectSingleNode("DisciplineStatistics") == null)
                                continue;

                            if (auto.InitialSummary.SelectSingleNode("DisciplineStatistics/Merit") != null)
                            {
                                if (auto.InitialMeritA + auto.InitialMeritB + auto.InitialMeritC > 0)
                                {
                                    IsValue = true;
                                }
                            }
                        }

                        if (IsValue)
                        {
                            //如果不是第一頁，就在上一頁的資料列下邊加黑線
                            if (index != 0)
                                ws.Cells.CreateRange(index - 1, 0, 1, columnNumber).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

                            //複製 Header
                            tool.CopyStyle(ws.Cells.CreateRange(index, 4, false), ptHeader);

                            dataIndex = index + 4;

                            //填寫基本資料
                            ws.Cells[index, 0].PutValue(TitleName1 + "(1/1)");
                            ws.Cells[index + 1, 0].PutValue(TitleName2);

                            #region 插入獎懲統計資訊
                            //獎懲統計資訊
                            Range disciplineStatisticsRange = ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber);
                            tool.CopyStyle(disciplineStatisticsRange, ptEachRow);
                            disciplineStatisticsRange.Merge();

                            //disciplineStatisticsRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Double, Color.Black);
                            //disciplineStatisticsRange.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Double, Color.Black);
                            disciplineStatisticsRange.RowHeight = 14.0;

                            tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                            tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                            tool.SetFontSize(ws.Cells[dataIndex, 0], 10);

                            ws.Cells[dataIndex, 0].PutValue("(無明細資料)");
                            dataIndex++;

                            #endregion

                            //獎懲統計資訊
                            disciplineStatisticsRange = ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber);
                            tool.CopyStyle(disciplineStatisticsRange, ptEachRow);

                            disciplineStatisticsRange.Merge();
                            //disciplineStatisticsRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Double, Color.Black);
                            disciplineStatisticsRange.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.None, Color.Transparent);
                            //disciplineStatisticsRange.RowHeight = 14.0;
                            
                            tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                            tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                            tool.SetFontSize(ws.Cells[dataIndex, 0], 10);

                            ws.Cells[dataIndex, 0].PutValue("獎勵非明細(僅轉入生與特殊狀況之學生會以非明細記錄)：");

                            foreach (AutoSummaryRecord auto in AutoSummaryDic[studentInfo.ID])
                            {
                                if (auto.InitialSummary == null)
                                    continue;

                                if (auto.InitialSummary.SelectSingleNode("DisciplineStatistics") == null)
                                    continue;

                                dataIndex++;
                                StringBuilder text = new StringBuilder("");
                                text.Append(auto.SchoolYear.ToString() + "/" + auto.Semester.ToString() + "　");

                                if (auto.InitialSummary.SelectSingleNode("DisciplineStatistics/Merit") != null)
                                {
                                    if (auto.InitialMeritA + auto.InitialMeritB + auto.InitialMeritC > 0)
                                    {
                                        if (auto.InitialMeritA > 0)
                                            text.Append("大功：" + auto.InitialMeritA + "　");
                                        if (auto.InitialMeritB > 0)
                                            text.Append("小功：" + auto.InitialMeritB + "　");
                                        if (auto.InitialMeritC > 0)
                                            text.Append("嘉獎：" + auto.InitialMeritC + "　");
                                    }
                                }

                                //獎懲統計內容
                                tool.CopyStyle(ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber), ptEachRow);

                                ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber).RowHeight = 14.0;
                                ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber).Merge();
                                ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber).SetOutlineBorder(BorderType.TopBorder, CellBorderType.None, Color.Transparent);
                                ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.None, Color.Transparent);
                                
                                tool.SetHorizontalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                                tool.SetVerticalAlignment(ws.Cells[dataIndex, 0], TextAlignmentType.Left);
                                tool.SetFontSize(ws.Cells[dataIndex, 0], 10);
                                tool.SetShrinkToFit(ws.Cells[dataIndex, 0], true);

                                ws.Cells[dataIndex, 0].PutValue(text.ToString());
                            }

                            ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                            ws.Cells.CreateRange(dataIndex, 0, 1, columnNumber).SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                            dataIndex++;
                        }
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
                        ws.HorizontalPageBreaks.Add(index, columnNumber);
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
                ws.Cells.CreateRange(dataIndex - 1, 0, 1, columnNumber).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
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

        public int SortByOccurDate(MeritRecord ASR1, MeritRecord ASR2)
        {
            return ASR1.OccurDate.CompareTo(ASR2.OccurDate);
        }
    }
}
