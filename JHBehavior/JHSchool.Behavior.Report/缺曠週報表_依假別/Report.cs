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

namespace JHSchool.Behavior.Report.缺曠週報表_依假別
{
    internal class Report : IReport
    {
        #region IReport 成員

        private BackgroundWorker _BGWAbsenceWeekListByAbsence;

        public void Print()
        {
            if (Class.Instance.SelectedList.Count == 0)
            {
                MsgBox.Show("未選擇學生!!");
                return;
            }

            WARCByAbsence weekForm = new WARCByAbsence();

            if (weekForm.ShowDialog() == DialogResult.OK)
            {
                Dictionary<string, List<string>> config = new Dictionary<string, List<string>>();

                //XmlElement preferenceData = CurrentUser.Instance.Preference["缺曠週報表_依缺曠別統計_列印設定"];
                ConfigData cd = User.Configuration["缺曠週報表_依缺曠別統計_列印設定"];
                XmlElement preferenceData = cd.GetXml("XmlData", null);

                if (preferenceData == null)
                {
                    MsgBox.Show("第一次使用缺曠週報表請先執行 列印設定 。");
                    return;
                }
                else
                {
                    foreach (XmlElement type in preferenceData.SelectNodes("Type"))
                    {
                        string prefix = type.GetAttribute("Text");
                        if (!config.ContainsKey(prefix))
                            config.Add(prefix, new List<string>());

                        foreach (XmlElement absence in type.SelectNodes("Absence"))
                        {
                            if (!config[prefix].Contains(absence.GetAttribute("Text")))
                                config[prefix].Add(absence.GetAttribute("Text"));
                        }
                    }
                }

                int CountAbsence = 0;
                foreach (string each in config.Keys)
                {
                    CountAbsence += config[each].Count;
                }
                if (CountAbsence == 0)
                {
                    MsgBox.Show("未選擇列印假別,請由 假別設定 進行選擇!");
                    return;
                }
                else
                {
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化缺曠週報表...");

                    object[] args = new object[] { config, weekForm.StartDate, weekForm.EndDate, weekForm.PaperSize, weekForm.ClassCix, weekForm.WeekCix };

                    _BGWAbsenceWeekListByAbsence = new BackgroundWorker();
                    _BGWAbsenceWeekListByAbsence.DoWork += new DoWorkEventHandler(_BGWAbsenceWeekListByAbsence_DoWork);
                    _BGWAbsenceWeekListByAbsence.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CommonMethods.ExcelReport_RunWorkerCompleted);
                    _BGWAbsenceWeekListByAbsence.ProgressChanged += new ProgressChangedEventHandler(CommonMethods.Report_ProgressChanged);
                    _BGWAbsenceWeekListByAbsence.WorkerReportsProgress = true;
                    _BGWAbsenceWeekListByAbsence.RunWorkerAsync(args);
                }
            }
        }

        #endregion

        #region 缺曠週報表_依缺曠別統計

        void _BGWAbsenceWeekListByAbsence_DoWork(object sender, DoWorkEventArgs e)
        {
            string reportName = "缺曠週報表";

            object[] args = e.Argument as object[];

            Dictionary<string, List<string>> config = args[0] as Dictionary<string, List<string>>;
            DateTime startDate = (DateTime)args[1];
            DateTime endDate = (DateTime)args[2];
            int size = (int)args[3];
            bool CheckClass = (bool)args[4];
            bool CheckWeek = (bool)args[5];

            DateTime firstDate = startDate;

            #region 快取學生缺曠紀錄資料

            List<ClassRecord> selectedClass = Class.Instance.SelectedList;
            selectedClass = SortClassIndex.JHSchool_ClassRecord(selectedClass);

            Dictionary<string, List<StudentRecord>> classStudentList = new Dictionary<string, List<StudentRecord>>();

            List<string> allStudentID = new List<string>();

            //紀錄每一個 Column 的 Index
            Dictionary<string, int> columnTable = new Dictionary<string, int>();

            //紀錄每一個學生的缺曠紀錄
            Dictionary<string, Dictionary<string, Dictionary<string, int>>> studentAbsenceList = new Dictionary<string, Dictionary<string, Dictionary<string, int>>>();

            //紀錄每一個學生本週累計的缺曠紀錄
            Dictionary<string, Dictionary<string, int>> studentWeekAbsenceList = new Dictionary<string, Dictionary<string, int>>();

            //紀錄每一個學生學期累計的缺曠紀錄
            Dictionary<string, Dictionary<string, int>> studentSemesterAbsenceList = new Dictionary<string, Dictionary<string, int>>();

            //節次對照表
            Dictionary<string, string> periodList = new Dictionary<string, string>();

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

            //取得 Period List
            DSResponse dsrsp = JHSchool.Compatibility.Feature.Basic.Config.GetPeriodList();
            foreach (XmlElement var in dsrsp.GetContent().GetElements("Period"))
            {
                string name = var.GetAttribute("Name");
                string type = var.GetAttribute("Type");
                if (!periodList.ContainsKey(name))
                    periodList.Add(name, type);
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
                    studentAbsenceList.Add(studentID, new Dictionary<string, Dictionary<string, int>>());
                if (!studentAbsenceList[studentID].ContainsKey(occurDate))
                    studentAbsenceList[studentID].Add(occurDate, new Dictionary<string, int>());

                if (!studentWeekAbsenceList.ContainsKey(studentID))
                    studentWeekAbsenceList.Add(studentID, new Dictionary<string, int>());

                foreach (XmlElement period in var.SelectNodes("Detail/Attendance/Period"))
                {
                    string type = periodList.ContainsKey(period.InnerText) ? periodList[period.InnerText] : "";
                    string absence = period.GetAttribute("AbsenceType");
                    if (!studentAbsenceList[studentID][occurDate].ContainsKey(type + "_" + absence))
                        studentAbsenceList[studentID][occurDate].Add(type + "_" + absence, 0);
                    studentAbsenceList[studentID][occurDate][type + "_" + absence]++;

                    if (!studentWeekAbsenceList[studentID].ContainsKey(type + "_" + absence))
                        studentWeekAbsenceList[studentID].Add(type + "_" + absence, 0);
                    studentWeekAbsenceList[studentID][type + "_" + absence]++;
                }
            }

            //產生 DSRequest，本學期累計
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
                    string type = periodList.ContainsKey(period.InnerText) ? periodList[period.InnerText] : "";
                    string absence = period.GetAttribute("AbsenceType");
                    if (!studentSemesterAbsenceList[studentID].ContainsKey(type + "_" + absence))
                        studentSemesterAbsenceList[studentID].Add(type + "_" + absence, 0);
                    studentSemesterAbsenceList[studentID][type + "_" + absence]++;
                }
            }

            #endregion

            //計算使用者自訂項目
            int allAbsenceNumber = 7;
            foreach (string type in config.Keys)
            {
                allAbsenceNumber += config[type].Count;
            }
            int current = 1;
            int all = allAbsenceNumber + allStudentNumber;

            #region 動態產生範本

            Workbook template = new Workbook();

            template.Open(new MemoryStream(ProjectResource.缺曠週報表_依假別), FileFormatType.Excel2003);

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

            //根據使用者設定的缺曠別，產生 Column

            List<string> perList = new List<string>();

            foreach (string type in config.Keys)
            {
                if (config[type].Count == 0)
                    continue;

                int typeStartIndex = colIndex;

                foreach (string var in config[type])
                {
                    string SaveVar = type + "_" + var;
                    prototype.Worksheets[0].Cells.CreateRange(colIndex, 1, true).Copy(tempEachColumn);
                    prototype.Worksheets[0].Cells[titleRow + 2, colIndex].PutValue(var);
                    columnTable.Add(SaveVar, colIndex - 3);

                    if(!perList.Contains(SaveVar))
                    {
                        perList.Add(type + "_" + var);
                    }

                    colIndex++;
                    _BGWAbsenceWeekListByAbsence.ReportProgress((int)(((double)current++ * 100.0) / (double)all));
                }

                int typeEndIndex = colIndex;

                Range typeRange = prototype.Worksheets[0].Cells.CreateRange(titleRow, typeStartIndex, titleRow + 2, typeEndIndex - typeStartIndex);
                typeRange.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Medium, Color.Black);
                typeRange.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);

                prototype.Worksheets[0].Cells.CreateRange(titleRow + 1, typeStartIndex, 1, typeEndIndex - typeStartIndex).Merge();
                prototype.Worksheets[0].Cells[titleRow + 1, typeStartIndex].PutValue(type);
            }

            dayEndIndex = colIndex;
            dayColumnNumber = dayEndIndex - dayStartIndex;
            if (dayColumnNumber == 0)
                dayColumnNumber = 1;

            prototype.Worksheets[0].Cells.CreateRange(titleRow, dayStartIndex, 1, dayColumnNumber).Merge();

            Range dayRange = prototype.Worksheets[0].Cells.CreateRange(dayStartIndex, dayColumnNumber, true);
            prototype.Worksheets[0].Cells[titleRow, dayStartIndex].PutValue(firstDate.ToShortDateString() + " (" + CommonMethods.GetChineseDayOfWeek(firstDate) + ")");
            columnTable.Add(firstDate.ToShortDateString(), dayStartIndex);

            //以一個日期為單位進行 Column 複製
            for (int i = 1; i < dayNumber; i++)
            {
                firstDate = firstDate.AddDays(1);

                dayStartIndex += dayColumnNumber;
                prototype.Worksheets[0].Cells.CreateRange(dayStartIndex, dayColumnNumber, true).Copy(dayRange);
                prototype.Worksheets[0].Cells[titleRow, dayStartIndex].PutValue(firstDate.ToShortDateString() + " (" + CommonMethods.GetChineseDayOfWeek(firstDate) + ")");
                columnTable.Add(firstDate.ToShortDateString(), dayStartIndex);
                _BGWAbsenceWeekListByAbsence.ReportProgress((int)(((double)current++ * 100.0) / (double)all));
            }

            dayStartIndex += dayColumnNumber;
            prototype.Worksheets[0].Cells.CreateRange(dayStartIndex, dayColumnNumber, true).Copy(dayRange);
            prototype.Worksheets[0].Cells[titleRow, dayStartIndex].PutValue("本週合計");
            columnTable.Add("本週合計", dayStartIndex);
            _BGWAbsenceWeekListByAbsence.ReportProgress((int)(((double)current++ * 100.0) / (double)all));

            dayStartIndex += dayColumnNumber;
            prototype.Worksheets[0].Cells.CreateRange(dayStartIndex, dayColumnNumber, true).Copy(dayRange);
            prototype.Worksheets[0].Cells[titleRow, dayStartIndex].PutValue("本學期累計");
            columnTable.Add("本學期累計", dayStartIndex);
            _BGWAbsenceWeekListByAbsence.ReportProgress((int)(((double)current++ * 100.0) / (double)all));

            dayStartIndex += dayColumnNumber;

            //合併標題列
            prototype.Worksheets[0].Cells.CreateRange(0, 0, 1, dayStartIndex).Merge();
            prototype.Worksheets[0].Cells.CreateRange(1, 0, 1, dayStartIndex).Merge();

            Range prototypeRow = prototype.Worksheets[0].Cells.CreateRange(5, 1, false);
            Range prototypeHeader = prototype.Worksheets[0].Cells.CreateRange(0, 5, false);

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
                            foreach (string ajn in studentWeekAbsenceList[CheckStudentData.ID].Keys) //取得該學生假別
                            {
                                if (perList.Contains(ajn)) //如果假別的確包含於清單中
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
                    ws.Cells.CreateRange(index, 5, false).Copy(prototypeHeader);

                    //填寫基本資料

                    string TeacherName = "";
                    if (classInfo.Teacher != null)
                    {
                        TeacherName = classInfo.Teacher.Name + " 老師";
                    }

                    if (CheckWeek) //new,True就是取得至星期日內
                    {
                        ws.Cells[index + 1, 0].PutValue("班級名稱： " + classInfo.Name + "　　班導師： " + TeacherName + "　　缺曠統計區間： " + startDate.ToShortDateString() + " ~ " + endDate.ToShortDateString());
                    }
                    else
                    {
                        ws.Cells[index + 1, 0].PutValue("班級名稱： " + classInfo.Name + "　　班導師： " + TeacherName + "　　缺曠統計區間： " + startDate.ToShortDateString() + " ~ " + endDate.AddDays(-2).ToShortDateString());
                    }
                    dataIndex = index + 5;

                    int studentCount = 0;
                    while (studentCount < classStudent.Count)
                    {
                        //複製每一個 row
                        ws.Cells.CreateRange(dataIndex, 1, false).Copy(prototypeRow);
                        if (studentCount % 5 == 0 && studentCount != 0)
                        {
                            Range eachFiveRow = ws.Cells.CreateRange(dataIndex, 0, 1, dayStartIndex);
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
                                Dictionary<string, int> dateAbsence = studentAbsenceList[studentID][date];

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
                            startCol = columnTable["本週合計"];

                            Dictionary<string, int> studentWeek = studentWeekAbsenceList[studentID];

                            foreach (string var in studentWeek.Keys)
                            {
                                if (columnTable.ContainsKey(var))
                                    ws.Cells[dataIndex, startCol + columnTable[var]].PutValue(studentWeekAbsenceList[studentID][var]);
                            }
                        }

                        if (studentSemesterAbsenceList.ContainsKey(studentID))
                        {
                            startCol = columnTable["本學期累計"];

                            Dictionary<string, int> studentSemester = studentSemesterAbsenceList[studentID];

                            foreach (string var in studentSemester.Keys)
                            {
                                if (columnTable.ContainsKey(var))
                                    ws.Cells[dataIndex, startCol + columnTable[var]].PutValue(studentSemester[var]);
                            }
                        }

                        studentCount++;
                        dataIndex++;
                        _BGWAbsenceWeekListByAbsence.ReportProgress((int)(((double)current++ * 100.0) / (double)all));
                    }

                    //資料列上邊各加上黑線
                    ws.Cells.CreateRange(index + 4, 0, 1, dayStartIndex).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);

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
