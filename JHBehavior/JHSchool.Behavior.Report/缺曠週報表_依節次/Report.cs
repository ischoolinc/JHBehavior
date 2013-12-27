using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using Aspose.Cells;
using FISCA.DSAUtil;
using Framework;

namespace JHSchool.Behavior.Report.缺曠週報表_依節次
{
    internal class Report : IReport
    {
        #region IReport 成員

        private BackgroundWorker _BGWAbsenceWeekListByPeriod;

        public void Print()
        {
            if (Class.Instance.SelectedList.Count == 0)
            {
                MsgBox.Show("未選擇學生!!");
                return;
            }

            List<string> config = new List<string>();

            WeekAbsenceReportCountByPeriod weekForm = new WeekAbsenceReportCountByPeriod();

            if (weekForm.ShowDialog() == DialogResult.OK)
            {
                //XmlElement preferenceData = CurrentUser.Instance.Preference["缺曠週報表_依節次統計_列印設定"];
                #region 節次
                ConfigData cd = User.Configuration[weekForm.ConfigName];
                XmlElement preferenceData = cd.GetXml("XmlData", null);

                if (preferenceData == null)
                {
                    MsgBox.Show("第一次使用缺曠週報表請先執行 節次設定。");
                    return;
                }
                else
                {
                    foreach (XmlElement period in preferenceData.SelectNodes("Period"))
                    {
                        string name = period.GetAttribute("Name");
                        config.Add(name);
                    }
                }

                if (config.Count == 0)
                {
                    MsgBox.Show("未選擇列印節次,請由 節次設定 進行選擇!");
                    return;
                }
                #endregion

                #region 假別

                Dictionary<string, List<string>> configDic = new Dictionary<string, List<string>>();

                cd = User.Configuration[weekForm.ConfigTypeName];
                preferenceData = cd.GetXml("XmlData", null);

                if (preferenceData == null)
                {
                    MsgBox.Show("第一次使用缺曠週報表請先執行 假別設定。");
                    return;
                }
                else
                {
                    foreach (XmlElement type in preferenceData.SelectNodes("Type"))
                    {
                        string prefix = type.GetAttribute("Text");
                        if (!configDic.ContainsKey(prefix))
                            configDic.Add(prefix, new List<string>());

                        foreach (XmlElement absence in type.SelectNodes("Absence"))
                        {
                            if (!configDic[prefix].Contains(absence.GetAttribute("Text")))
                                configDic[prefix].Add(absence.GetAttribute("Text"));
                        }
                    }
                }

                int CountAbsence = 0;
                foreach (string each in configDic.Keys)
                {
                    CountAbsence += configDic[each].Count;
                }

                if (configDic.Count <= 0)
                {
                    MsgBox.Show("假別設定 必須勾選至少一個假別。");
                    return;
                }
                #endregion

                FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化缺曠週報表...");

                object[] args = new object[] { config, configDic, weekForm.StartDate, weekForm.EndDate, weekForm.PaperSize, weekForm.ClassCix, weekForm.WeekCix }; //ClassCix增加的!!

                _BGWAbsenceWeekListByPeriod = new BackgroundWorker();
                _BGWAbsenceWeekListByPeriod.DoWork += new DoWorkEventHandler(_BGWAbsenceWeekListByPeriod_DoWork);
                _BGWAbsenceWeekListByPeriod.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CommonMethods.ExcelReport_RunWorkerCompleted);
                _BGWAbsenceWeekListByPeriod.ProgressChanged += new ProgressChangedEventHandler(CommonMethods.Report_ProgressChanged);
                _BGWAbsenceWeekListByPeriod.WorkerReportsProgress = true;
                _BGWAbsenceWeekListByPeriod.RunWorkerAsync(args);

            }
        }

        #endregion

        #region 缺曠週報表_依節次統計

        void _BGWAbsenceWeekListByPeriod_DoWork(object sender, DoWorkEventArgs e)
        {
            string reportName = "缺曠週報表";

            object[] args = e.Argument as object[];

            List<string> config = args[0] as List<string>;
            Dictionary<string, List<string>> configDic = args[1] as Dictionary<string, List<string>>;
            List<string> AbsenceList = new List<string>();
            foreach (string each1 in configDic.Keys)
            {
                foreach (string each2 in configDic[each1])
                {
                    if (!AbsenceList.Contains(each2))
                    {
                        AbsenceList.Add(each2);
                    }
                }
            }
            DateTime startDate = (DateTime)args[2];
            DateTime endDate = (DateTime)args[3];
            int size = (int)args[4];
            bool CheckClass = (bool)args[5]; //增加的!!
            bool CheckWeek = (bool)args[6]; //增加的!!
            DateTime firstDate = startDate;

            #region 快取學生缺曠紀錄資料

            List<ClassRecord> selectedClass = Class.Instance.SelectedList;
            selectedClass = SortClassIndex.JHSchool_ClassRecord(selectedClass);

            Dictionary<string, List<StudentRecord>> classStudentList = new Dictionary<string, List<StudentRecord>>();

            List<string> allStudentID = new List<string>();

            //紀錄每一個 Column 的 Index
            Dictionary<string, int> columnTable = new Dictionary<string, int>();

            //記錄缺曠類別 Column 的 Index
            Dictionary<string, int> columnTable_asbs = new Dictionary<string, int>();

            //紀錄每一個學生的缺曠紀錄
            Dictionary<string, Dictionary<string, Dictionary<string, string>>> studentAbsenceList = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();

            //紀錄每一個學生本週累計的缺曠紀錄
            Dictionary<string, Dictionary<string, int>> studentWeekAbsenceList = new Dictionary<string, Dictionary<string, int>>();

            //紀錄每一個學生學期累計的缺曠紀錄
            Dictionary<string, Dictionary<string, int>> studentSemesterAbsenceList = new Dictionary<string, Dictionary<string, int>>();

            //使用者定義的節次列表
            List<string> userDefinedPeriodList = new List<string>();

            //節次對照表
            List<string> periodList = new List<string>();

            //缺曠對照表
            Dictionary<string, string> absenceList = new Dictionary<string, string>();

            int allStudentNumber = 0;

            //計算學生總數，取得所有學生ID
            foreach (ClassRecord aClass in selectedClass)
            {
                List<StudentRecord> classStudent = aClass.Students.GetInSchoolStudents(); //取得在校生

                //List<StudentRecord> classStudent = new List<StudentRecord>(); //取得一般生
                //foreach (StudentRecord each in aClass.Students)
                //{
                //    if (each.Status == "一般")
                //    {
                //        classStudent.Add(each);
                //    }
                //}

                classStudent.Sort(new Comparison<StudentRecord>(CommonMethods.ClassSeatNoComparer));

                foreach (StudentRecord aStudent in classStudent)
                {
                    allStudentID.Add(aStudent.ID);
                }
                if (!classStudentList.ContainsKey(aClass.ID))
                    classStudentList.Add(aClass.ID, classStudent);
                allStudentNumber += classStudent.Count;
            }

            //套用使用者的設定
            userDefinedPeriodList = config;

            //取得 Period List
            DSResponse dsrsp = JHSchool.Compatibility.Feature.Basic.Config.GetPeriodList();
            Dictionary<string, string> PeriodTypeDic = new Dictionary<string, string>();
            foreach (XmlElement var in dsrsp.GetContent().GetElements("Period"))
            {
                if (!periodList.Contains(var.GetAttribute("Name")))
                    periodList.Add(var.GetAttribute("Name"));
                if (!PeriodTypeDic.ContainsKey(var.GetAttribute("Name")))
                {
                    PeriodTypeDic.Add(var.GetAttribute("Name"), var.GetAttribute("Type"));
                }
            }

            //取得 Absence List
            dsrsp = JHSchool.Compatibility.Feature.Basic.Config.GetAbsenceList();
            foreach (XmlElement var in dsrsp.GetContent().GetElements("Absence"))
            {
                if (!absenceList.ContainsKey(var.GetAttribute("Name")))
                    absenceList.Add(var.GetAttribute("Name"), var.GetAttribute("Abbreviation"));
            }

            //產生 DSRequest
            DSXmlHelper helper = new DSXmlHelper("Request");
            helper.AddElement("Field");
            helper.AddElement("Field", "All");
            helper.AddElement("Condition");
            foreach (string var in allStudentID)
            {
                helper.AddElement("Condition", "RefStudentID", var);
            }
            helper.AddElement("Condition", "StartDate", startDate.ToShortDateString());

            if (CheckWeek) //new,True就是取得至星期日內
            {
                helper.AddElement("Condition", "EndDate", endDate.ToShortDateString());
            }
            else //new,false就是取得到星期五的缺曠內容
            {
                helper.AddElement("Condition", "EndDate", endDate.AddDays(-2).ToShortDateString());
            }

            helper.AddElement("Order");
            helper.AddElement("Order", "OccurDate", "desc");
            dsrsp = JHSchool.Compatibility.Feature.Student.QueryAttendance.GetAttendance(new DSRequest(helper));

            foreach (XmlElement var in dsrsp.GetContent().GetElements("Attendance"))
            {
                string studentID = var.SelectSingleNode("RefStudentID").InnerText;
                string occurDate = DateTime.Parse(var.SelectSingleNode("OccurDate").InnerText).ToShortDateString();

                if (!studentAbsenceList.ContainsKey(studentID))
                    studentAbsenceList.Add(studentID, new Dictionary<string, Dictionary<string, string>>());
                if (!studentAbsenceList[studentID].ContainsKey(occurDate))
                    studentAbsenceList[studentID].Add(occurDate, new Dictionary<string, string>());

                if (!studentWeekAbsenceList.ContainsKey(studentID))
                    studentWeekAbsenceList.Add(studentID, new Dictionary<string, int>());

                foreach (XmlElement period in var.SelectNodes("Detail/Attendance/Period"))
                {
                    string innertext = period.InnerText;
                    string absence = period.GetAttribute("AbsenceType");

                    if (PeriodTypeDic.ContainsKey(innertext))
                    {
                        if (configDic.ContainsKey(PeriodTypeDic[innertext]))
                        {
                            if (configDic[PeriodTypeDic[innertext]].Contains(absence))
                            {
                                if (!studentAbsenceList[studentID][occurDate].ContainsKey(innertext))
                                {
                                    if (absenceList.ContainsKey(absence))
                                    {
                                        studentAbsenceList[studentID][occurDate].Add(innertext, absenceList[absence]);
                                        if (!studentWeekAbsenceList[studentID].ContainsKey(absence))
                                            studentWeekAbsenceList[studentID].Add(absence, 0);

                                        //學生缺曠紀錄中的節次必須存在於對照表之中，才列入統計
                                        if (periodList.Contains(innertext))
                                            studentWeekAbsenceList[studentID][absence]++;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            //產生 DSRequest for 本學期累計
            helper = new DSXmlHelper("Request");
            helper.AddElement("Field");
            helper.AddElement("Field", "All");
            helper.AddElement("Condition");
            foreach (string var in allStudentID)
            {
                helper.AddElement("Condition", "RefStudentID", var);
            }
            helper.AddElement("Condition", "SchoolYear", School.DefaultSchoolYear);
            helper.AddElement("Condition", "Semester", School.DefaultSemester);

            if (CheckWeek) //new,True就是取得至星期日內
            {
                helper.AddElement("Condition", "EndDate", endDate.ToShortDateString());
            }
            else
            {
                helper.AddElement("Condition", "EndDate", endDate.AddDays(-2).ToShortDateString());
            }
            helper.AddElement("Order");
            helper.AddElement("Order", "OccurDate", "desc");
            dsrsp = JHSchool.Compatibility.Feature.Student.QueryAttendance.GetAttendance(new DSRequest(helper));

            foreach (XmlElement var in dsrsp.GetContent().GetElements("Attendance"))
            {
                string studentID = var.SelectSingleNode("RefStudentID").InnerText;

                if (!studentSemesterAbsenceList.ContainsKey(studentID))
                    studentSemesterAbsenceList.Add(studentID, new Dictionary<string, int>());

                foreach (XmlElement period in var.SelectNodes("Detail/Attendance/Period"))
                {
                    string innertext = period.InnerText;
                    string absence = period.GetAttribute("AbsenceType");

                    if (PeriodTypeDic.ContainsKey(innertext))
                    {
                        if (configDic.ContainsKey(PeriodTypeDic[innertext]))
                        {
                            if (configDic[PeriodTypeDic[innertext]].Contains(absence))
                            {
                                if (absenceList.ContainsKey(absence))
                                {
                                    if (!studentSemesterAbsenceList[studentID].ContainsKey(absence))
                                        studentSemesterAbsenceList[studentID].Add(absence, 0);

                                    //學生缺曠紀錄中的節次必須存在於對照表之中，才列入統計
                                    if (periodList.Contains(innertext))
                                        studentSemesterAbsenceList[studentID][absence]++;
                                }
                            }
                        }
                    }
                }
            }

            #endregion

            //計算使用者自訂項目
            int allAbsenceNumber = 6;
            allAbsenceNumber += userDefinedPeriodList.Count;
            allAbsenceNumber += absenceList.Count;
            int current = 1;
            int all = allAbsenceNumber + allStudentNumber;

            #region 動態產生範本

            Workbook template = new Workbook();

            template.Open(new MemoryStream(ProjectResource.缺曠週報表_依節次), FileFormatType.Excel2003);

            template.Worksheets[0].Cells.DeleteRow(3);
            template.Worksheets[0].Cells.CreateRange(2, 3, 2, 1).SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Medium, Color.Black);

            Range tempStudent = template.Worksheets[0].Cells.CreateRange(0, 3, true);
            Range tempEachColumn = template.Worksheets[0].Cells.CreateRange(3, 1, true);

            Workbook prototype = new Workbook();
            prototype.Copy(template);

            prototype.Worksheets[0].Cells.CreateRange(0, 3, true).Copy(tempStudent);

            int titleRow = 2;
            int dayNumber;
            if (CheckWeek)
            {
                dayNumber = 7;
            }
            else
            {
                dayNumber = 5;
            }

            int colIndex = 3;

            int dayStartIndex = colIndex;
            int dayEndIndex;
            int dayColumnNumber;

            //根據使用者定義的節次動態產生欄位
            foreach (string period in userDefinedPeriodList)
            {
                prototype.Worksheets[0].Cells.CreateRange(colIndex, 1, true).Copy(tempEachColumn);
                prototype.Worksheets[0].Cells[titleRow + 1, colIndex].PutValue(period);
                columnTable.Add(period, colIndex - 3);
                colIndex++;
                _BGWAbsenceWeekListByPeriod.ReportProgress((int)(((double)current++ * 100.0) / (double)all));
            }

            dayEndIndex = colIndex;
            dayColumnNumber = dayEndIndex - dayStartIndex;
            if (dayColumnNumber == 0)
                dayColumnNumber = 1;

            prototype.Worksheets[0].Cells.CreateRange(titleRow, dayStartIndex, 1, dayColumnNumber).Merge();

            Range dayRange = prototype.Worksheets[0].Cells.CreateRange(dayStartIndex, dayColumnNumber, true);
            prototype.Worksheets[0].Cells[titleRow, dayStartIndex].PutValue(firstDate.ToShortDateString() + " (" + CommonMethods.GetChineseDayOfWeek(firstDate) + ")");
            columnTable.Add(firstDate.ToShortDateString(), dayStartIndex);

            //設定每一日期的欄位外框
            prototype.Worksheets[0].Cells.CreateRange(titleRow, dayStartIndex, 3, dayColumnNumber).SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Medium, Color.Black);

            for (int i = 1; i < dayNumber; i++)
            {
                firstDate = firstDate.AddDays(1);

                dayStartIndex += dayColumnNumber;
                prototype.Worksheets[0].Cells.CreateRange(dayStartIndex, dayColumnNumber, true).Copy(dayRange);
                prototype.Worksheets[0].Cells[titleRow, dayStartIndex].PutValue(firstDate.ToShortDateString() + " (" + CommonMethods.GetChineseDayOfWeek(firstDate) + ")");
                columnTable.Add(firstDate.ToShortDateString(), dayStartIndex);
                _BGWAbsenceWeekListByPeriod.ReportProgress((int)(((double)current++ * 100.0) / (double)all));
            }

            dayStartIndex += dayColumnNumber;

            colIndex = dayStartIndex;

            //產生所有缺曠別的欄位
            foreach (string type in AbsenceList)
            {
                prototype.Worksheets[0].Cells.CreateRange(colIndex, 1, true).Copy(tempEachColumn);
                prototype.Worksheets[0].Cells[titleRow + 1, colIndex].PutValue(type);
                columnTable_asbs.Add(type, colIndex - dayStartIndex); //缺曠別的Index另外記錄,以免錯誤
                //columnTable.Add(type, colIndex - dayStartIndex);
                colIndex++;
                _BGWAbsenceWeekListByPeriod.ReportProgress((int)(((double)current++ * 100.0) / (double)all));
            }

            prototype.Worksheets[0].Cells.CreateRange(titleRow, dayStartIndex, 3, AbsenceList.Count).SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Medium, Color.Black);

            prototype.Worksheets[0].Cells.CreateRange(titleRow, dayStartIndex, 1, AbsenceList.Count).Merge();
            Range weekCountRange = prototype.Worksheets[0].Cells.CreateRange(dayStartIndex, AbsenceList.Count, true);
            prototype.Worksheets[0].Cells.CreateRange(dayStartIndex + AbsenceList.Count, AbsenceList.Count, true).Copy(weekCountRange);
            _BGWAbsenceWeekListByPeriod.ReportProgress((int)(((double)current++ * 100.0) / (double)all));

            prototype.Worksheets[0].Cells[titleRow, dayStartIndex].PutValue("本週合計");
            columnTable.Add("本週合計", dayStartIndex);

            dayStartIndex += AbsenceList.Count;

            prototype.Worksheets[0].Cells.CreateRange(titleRow, dayStartIndex, 3, AbsenceList.Count).SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Medium, Color.Black);

            prototype.Worksheets[0].Cells[titleRow, dayStartIndex].PutValue("本學期累計");
            columnTable.Add("本學期累計", dayStartIndex);

            dayStartIndex += AbsenceList.Count;

            prototype.Worksheets[0].Cells.CreateRange(0, 0, 1, dayStartIndex).Merge();
            prototype.Worksheets[0].Cells.CreateRange(1, 0, 1, dayStartIndex).Merge();

            Range prototypeRow = prototype.Worksheets[0].Cells.CreateRange(4, 1, false);
            Range prototypeHeader = prototype.Worksheets[0].Cells.CreateRange(0, 4, false);

            current++;

            #endregion

            #region 產生報表

            Workbook wb = new Workbook();
            wb.Copy(prototype);
            Worksheet ws = wb.Worksheets[0];

            #region 判斷紙張大小
            if (size == 0)
            {
                ws.PageSetup.PaperSize = PaperSizeType.PaperA3;
                ws.PageSetup.Zoom = 90;
            }
            else if (size == 1)
            {
                ws.PageSetup.PaperSize = PaperSizeType.PaperA4;
                ws.PageSetup.Zoom = 65;
            }
            else if (size == 2)
            {
                ws.PageSetup.PaperSize = PaperSizeType.PaperB4;
                ws.PageSetup.Zoom = 80;
            }
            #endregion

            int index = 0;
            int dataIndex = 0;

            List<string> list = new List<string>();

            #region 檢查是否有資料
            if (CheckClass) //如果需要過慮資料
            {
                foreach (ClassRecord CheckClassData in selectedClass)
                {
                    List<StudentRecord> classStudent = classStudentList[CheckClassData.ID];
                    bool jumpNext = false;

                    foreach (StudentRecord CheckStudentData in classStudent)
                    {
                        if (studentWeekAbsenceList.ContainsKey(CheckStudentData.ID)) //如果studentWeekAbsenceList內包含了該學生ID
                        {
                            if (!list.Contains(CheckClassData.ID)) //如果清單中不包含就加入
                            {
                                list.Add(CheckClassData.ID);
                                jumpNext = true;
                            }
                        }

                        if (jumpNext)
                            break;
                    }
                }
            }
            else //如果不需要過慮資料
            {
                foreach (ClassRecord CheckClassData in selectedClass)
                {
                    list.Add(CheckClassData.ID);
                }
            }
            #endregion

            foreach (ClassRecord classInfo in selectedClass)
            {
                if (list.Contains(classInfo.ID))
                {
                    List<StudentRecord> classStudent = classStudentList[classInfo.ID];

                    //如果不是第一頁，就在上一頁的資料列下邊加黑線
                    if (index != 0)
                        ws.Cells.CreateRange(index - 1, 0, 1, dayStartIndex).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);

                    //複製 Header
                    ws.Cells.CreateRange(index, 4, false).Copy(prototypeHeader);

                    //填寫基本資料
                    string TeacherName = "";
                    if (classInfo.Teacher != null)
                    {
                        TeacherName = classInfo.Teacher.Name + " 老師";
                    }

                    ws.Cells[index, 0].PutValue(School.DefaultSchoolYear + " 學年度 " + School.DefaultSemester + " 學期 " + School.ChineseName + " 缺曠週報表");
                    if (CheckWeek) //new,True就是取得至星期日內
                    {
                        ws.Cells[index + 1, 0].PutValue("班級名稱： " + classInfo.Name + "　　班導師： " + TeacherName + "　　缺曠統計區間： " + startDate.ToShortDateString() + " ~ " + endDate.ToShortDateString());
                    }
                    else
                    {
                        ws.Cells[index + 1, 0].PutValue("班級名稱： " + classInfo.Name + "　　班導師： " + TeacherName + "　　缺曠統計區間： " + startDate.ToShortDateString() + " ~ " + endDate.AddDays(-2).ToShortDateString());
                    }
                    dataIndex = index + 4;

                    int studentCount = 0;
                    while (studentCount < classStudent.Count)
                    {
                        //複製每一個 row
                        ws.Cells.CreateRange(dataIndex, 1, false).Copy(prototypeRow);
                        if (studentCount % 5 == 0 && studentCount != 0)
                        {
                            Range eachFiveRow = wb.Worksheets[0].Cells.CreateRange(dataIndex, 0, 1, dayStartIndex);
                            eachFiveRow.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Double, Color.Black);
                        }

                        //填寫學生缺曠資料
                        StudentRecord student = classStudent[studentCount];
                        string studentID = student.ID;
                        ws.Cells[dataIndex, 0].PutValue(student.StudentNumber);
                        ws.Cells[dataIndex, 1].PutValue(student.SeatNo);
                        ws.Cells[dataIndex, 2].PutValue(student.Name);

                        int startCol;
                        if (studentAbsenceList.ContainsKey(studentID))
                        {
                            foreach (string date in studentAbsenceList[studentID].Keys)
                            {
                                Dictionary<string, string> dateAbsence = studentAbsenceList[studentID][date];

                                startCol = columnTable[date];

                                foreach (string var in dateAbsence.Keys)
                                {
                                    if (columnTable.ContainsKey(var))
                                    {
                                        ws.Cells[dataIndex, startCol + columnTable[var]].PutValue(dateAbsence[var]);
                                    }
                                }
                            }
                        }

                        if (studentWeekAbsenceList.ContainsKey(studentID))
                        {
                            //寫入本週合計
                            startCol = columnTable["本週合計"];

                            Dictionary<string, int> studentWeek = studentWeekAbsenceList[studentID];

                            foreach (string var in studentWeek.Keys)
                            {
                                ws.Cells[dataIndex, startCol + columnTable_asbs[var]].PutValue(studentWeek[var]);
                            }
                        }

                        if (studentSemesterAbsenceList.ContainsKey(studentID))
                        {
                            //寫入學期累計
                            startCol = columnTable["本學期累計"];

                            Dictionary<string, int> studentSemester = studentSemesterAbsenceList[studentID];

                            foreach (string var in studentSemester.Keys)
                            {
                                if (columnTable_asbs.ContainsKey(var))
                                    ws.Cells[dataIndex, startCol + columnTable_asbs[var]].PutValue(studentSemester[var]);
                            }
                        }

                        studentCount++;
                        dataIndex++;
                        _BGWAbsenceWeekListByPeriod.ReportProgress((int)(((double)current++ * 100.0) / (double)all));
                    }

                    //資料列上邊加上黑線
                    ws.Cells.CreateRange(index + 3, 0, 1, dayStartIndex).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);

                    //表格最右邊加上黑線
                    ws.Cells.CreateRange(index + 2, dayStartIndex - 1, studentCount + 2, 1).SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);

                    index = dataIndex;

                    //設定分頁
                    ws.HPageBreaks.Add(index, dayStartIndex);
                }
            }

            //最後一頁的資料列下邊加上黑線
            if (dataIndex != 0)
            {
                ws.Cells.CreateRange(dataIndex - 1, 0, 1, dayStartIndex).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);
            }

            #endregion

            string path = Path.Combine(Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".xlt");
            e.Result = new object[] { reportName, path, wb };
        }
        #endregion

    }
}
