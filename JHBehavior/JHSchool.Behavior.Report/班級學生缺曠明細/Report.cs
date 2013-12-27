using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using Aspose.Cells;
using FISCA.DSAUtil;

namespace JHSchool.Behavior.Report.班級學生缺曠明細
{
    internal class Report : IReport
    {
        #region IReport 成員

        private BackgroundWorker _BGWClassStudentAbsenceDetail;

        public void Print()
        {
            if (Class.Instance.SelectedList.Count == 0)
                return;

            AbsenceDetailForm form = new AbsenceDetailForm();
            if (form.ShowDialog() == DialogResult.OK)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化班級學生缺曠明細...");

                object[] args = new object[] { form.StartDate, form.EndDate, form.PaperSize };

                _BGWClassStudentAbsenceDetail = new BackgroundWorker();
                _BGWClassStudentAbsenceDetail.DoWork += new DoWorkEventHandler(_BGWClassStudentAbsenceDetail_DoWork);
                _BGWClassStudentAbsenceDetail.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CommonMethods.ExcelReport_RunWorkerCompleted);
                _BGWClassStudentAbsenceDetail.ProgressChanged += new ProgressChangedEventHandler(CommonMethods.Report_ProgressChanged);
                _BGWClassStudentAbsenceDetail.WorkerReportsProgress = true;
                _BGWClassStudentAbsenceDetail.RunWorkerAsync(args);
            }
        }

        #endregion

        #region 班級學生缺曠明細

        void _BGWClassStudentAbsenceDetail_DoWork(object sender, DoWorkEventArgs e)
        {
            string reportName = "班級缺曠記錄明細";

            object[] args = e.Argument as object[];

            DateTime startDate = (DateTime)args[0];
            DateTime endDate = (DateTime)args[1];
            int size = (int)args[2];

            #region 快取資料

            List<ClassRecord> selectedClass = Class.Instance.SelectedList;
            selectedClass = SortClassIndex.JHSchool_ClassRecord(selectedClass);

            //學生ID查詢班級ID
            Dictionary<string, string> studentClassDict = new Dictionary<string, string>();

            //學生ID查詢學生資訊
            Dictionary<string, StudentRecord> studentInfoDict = new Dictionary<string, StudentRecord>();

            //缺曠明細，Key為 ClassID -> StudentID -> OccurDate -> Period
            Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, string>>>> allAbsenceDetail = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, string>>>>();

            //所有學生ID
            List<string> allStudentID = new List<string>();

            //缺曠筆數
            int currentCount = 1;
            int totalNumber = 0;

            //紀錄每一個 Column 的 Index
            Dictionary<string, int> columnTable = new Dictionary<string, int>();

            //節次對照表
            List<string> periodList = new List<string>();

            //建立學生班級對照表
            foreach (ClassRecord aClass in selectedClass)
            {
                List<StudentRecord> classStudent = aClass.Students;

                foreach (StudentRecord student in classStudent)
                {
                    if (student.Status == "一般" || student.Status == "輟學")
                    {
                        allStudentID.Add(student.ID);
                        studentClassDict.Add(student.ID, aClass.ID);
                        studentInfoDict.Add(student.ID, student);
                    }
                }

                allAbsenceDetail.Add(aClass.ID, new Dictionary<string, Dictionary<string, Dictionary<string, string>>>());
            }

            //取得 Period List
            DSResponse dsrsp = JHSchool.Compatibility.Feature.Basic.Config.GetPeriodList();
            foreach (XmlElement var in dsrsp.GetContent().GetElements("Period"))
            {
                if (!periodList.Contains(var.GetAttribute("Name")))
                    periodList.Add(var.GetAttribute("Name"));
            }

            //取得缺曠明細，產生 DSRequest
            DSXmlHelper helper = new DSXmlHelper("Request");
            helper.AddElement("Field");
            helper.AddElement("Field", "All");
            helper.AddElement("Condition");
            foreach (string var in allStudentID)
            {
                helper.AddElement("Condition", "RefStudentID", var);
            }
            helper.AddElement("Condition", "StartDate", startDate.ToShortDateString());
            helper.AddElement("Condition", "EndDate", endDate.ToShortDateString());
            helper.AddElement("Order");
            helper.AddElement("Order", "OccurDate", "asc");
            dsrsp = JHSchool.Compatibility.Feature.Student.QueryAttendance.GetAttendance(new DSRequest(helper));

            foreach (XmlElement var in dsrsp.GetContent().GetElements("Attendance"))
            {
                string studentID = var.SelectSingleNode("RefStudentID").InnerText;
                string occurDate = DateTime.Parse(var.SelectSingleNode("OccurDate").InnerText).ToShortDateString();
                string classID = studentClassDict[studentID];

                if (!allAbsenceDetail.ContainsKey(classID))
                    allAbsenceDetail.Add(classID, new Dictionary<string, Dictionary<string, Dictionary<string, string>>>());

                if (!allAbsenceDetail[classID].ContainsKey(studentID))
                    allAbsenceDetail[classID].Add(studentID, new Dictionary<string, Dictionary<string, string>>());

                if (!allAbsenceDetail[classID][studentID].ContainsKey(occurDate))
                    allAbsenceDetail[classID][studentID].Add(occurDate, new Dictionary<string, string>());

                foreach (XmlElement period in var.SelectNodes("Detail/Attendance/Period"))
                {
                    //節次
                    string innertext = period.InnerText;
                    //缺曠別
                    string absence = period.GetAttribute("AbsenceType");

                    if (!allAbsenceDetail[classID][studentID][occurDate].ContainsKey(innertext))
                        allAbsenceDetail[classID][studentID][occurDate].Add(innertext, absence);
                }

                //累計筆數
                totalNumber++;
            }

            #endregion

            #region 產生範本

            Workbook template = new Workbook();
            template.Open(new MemoryStream(ProjectResource.班級學生缺曠明細), FileFormatType.Excel2003);

            Range tempStudent = template.Worksheets[0].Cells.CreateRange(0, 3, true);
            Range tempEachColumn = template.Worksheets[0].Cells.CreateRange(3, 1, true);

            Workbook prototype = new Workbook();
            prototype.Copy(template);

            prototype.Worksheets[0].Cells.CreateRange(0, 3, true).Copy(tempStudent);

            int colIndex = 3;

            int startIndex = colIndex;
            int endIndex;
            int columnNumber;

            int splitLineIndex = 0;

            //根據節次對照表產生欄位，並且隨著節次字數調整欄寬
            foreach (string period in periodList)
            {
                Range each = prototype.Worksheets[0].Cells.CreateRange(colIndex, 1, true);
                each.Copy(tempEachColumn);
                if (period.Length > 2)
                {
                    double width = 4.5;
                    width += (double)((period.Length - 2) * 2);
                    each.ColumnWidth = width;
                }
                prototype.Worksheets[0].Cells[1, colIndex].PutValue(period);
                columnTable.Add(period, colIndex - 3);
                colIndex++;
            }

            endIndex = colIndex;
            splitLineIndex = (colIndex + 1) / 2;

            columnNumber = endIndex - startIndex;
            if (columnNumber == 0)
                columnNumber = 1;

            prototype.Worksheets[0].Cells.CreateRange(0, 0, 1, splitLineIndex).Merge();
            prototype.Worksheets[0].Cells.CreateRange(0, splitLineIndex, 1, endIndex - splitLineIndex).Merge();

            Range prototypeRow = prototype.Worksheets[0].Cells.CreateRange(2, 1, false);
            Range prototypeHeader = prototype.Worksheets[0].Cells.CreateRange(0, 2, false);

            #endregion

            #region 產生報表

            Workbook wb = new Workbook();
            wb.Copy(prototype);
            Worksheet ws = wb.Worksheets[0];

            #region 判斷紙張大小
            if (size == 0)
            {
                ws.PageSetup.PaperSize = PaperSizeType.PaperA3;
                ws.PageSetup.Zoom = 145;
            }
            else if (size == 1)
            {
                ws.PageSetup.PaperSize = PaperSizeType.PaperA4;
                ws.PageSetup.Zoom = 100;
            }
            else if (size == 2)
            {
                ws.PageSetup.PaperSize = PaperSizeType.PaperB4;
                ws.PageSetup.Zoom = 125;
            }
            #endregion

            int index = 0;
            int dataIndex = 0;

            foreach (ClassRecord classInfo in selectedClass)
            {
                string TitleName1 = classInfo.Name + " 班級缺曠記錄明細";
                string TitleName2 = "統計日期：" + startDate.ToShortDateString() + " ~ " + endDate.ToShortDateString();
            
                //班級/學生/日期/詳細資料
                Dictionary<string, Dictionary<string, Dictionary<string, string>>> classAbsenceDetail = allAbsenceDetail[classInfo.ID];

                if (classAbsenceDetail.Count <= 0)
                    continue;

                //如果不是第一頁，就在上一頁的資料列下邊加黑線
                if (index != 0)
                    ws.Cells.CreateRange(index - 1, 0, 1, endIndex).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

                //複製 Header
                wb.Worksheets[0].Cells.CreateRange(index, 2, false).Copy(prototypeHeader);

                dataIndex = index + 2;

                //學生Row筆數超過40筆,則添加換行符號,與標頭
                int CountRows = 0;

                //取總頁數
                int TotlePage = 0;
                //資料筆數count
                int CountDetail = 0;
                foreach(string each1 in classAbsenceDetail.Keys) //班級
                {
                    CountDetail += classAbsenceDetail[each1].Count;
                }
                TotlePage = CountDetail / 40; //取總頁數

                int pageCount = 1; //第一頁
                if (CountDetail % 40 != 0)
                {
                    TotlePage++; //如果有餘數代表要加一頁
                }

                //填寫基本資料
                ws.Cells[index, 0].PutValue(TitleName1 + "(" + pageCount.ToString() + "/" + TotlePage.ToString() + ")");
                pageCount++;
                ws.Cells[index, splitLineIndex].PutValue(TitleName2);
                ws.Cells[index, splitLineIndex].Style.Font.Size = 10;
                ws.Cells[index, splitLineIndex].Style.HorizontalAlignment = TextAlignmentType.Left;
                ws.Cells[index, splitLineIndex].Style.VerticalAlignment = TextAlignmentType.Bottom;

                List<string> list = new List<string>(classAbsenceDetail.Keys);
                list.Sort(new SeatNoComparer(studentInfoDict));

                foreach (string studentID in list)
                {
                    //填寫資料
                    Dictionary<string, Dictionary<string, string>> studentAbsenceDetail = classAbsenceDetail[studentID];
                    foreach (string occurDate in studentAbsenceDetail.Keys)
                    {
                        //計算筆數
                        CountRows++;

                        //先檢查他的缺曠節次是否都不存在於對照表中，不存在則不印該筆資料
                        bool printable = false;
                        foreach (string period in studentAbsenceDetail[occurDate].Keys)
                        {
                            if (columnTable.ContainsKey(period))
                                printable = true;
                        }
                        if (!printable)
                            continue;

                        //複製每一個 row
                        ws.Cells.CreateRange(dataIndex, 1, false).Copy(prototypeRow);

                        ws.Cells[dataIndex, 0].PutValue(studentInfoDict[studentID].SeatNo);
                        ws.Cells[dataIndex, 1].PutValue(studentInfoDict[studentID].Name);
                        ws.Cells[dataIndex, 2].PutValue(occurDate);

                        foreach (string period in studentAbsenceDetail[occurDate].Keys)
                        {
                            if (columnTable.ContainsKey(period))
                                ws.Cells[dataIndex, columnTable[period] + 3].PutValue(studentAbsenceDetail[occurDate][period]);
                        }

                        dataIndex++;
                        _BGWClassStudentAbsenceDetail.ReportProgress((int)(((double)currentCount++ * 100.0) / (double)totalNumber));

                        if (CountRows == 40 && pageCount <= TotlePage)
                        {
                            CountRows = 0;
                            //分頁
                            ws.HPageBreaks.Add(dataIndex, endIndex);
                            //複製 Header
                            ws.Cells.CreateRange(dataIndex, 2, false).Copy(prototypeHeader);

                            ws.Cells[dataIndex, 0].PutValue(TitleName1 + "(" + pageCount.ToString() + "/" + TotlePage.ToString() + ")");
                            pageCount++;
                            ws.Cells[dataIndex, splitLineIndex].PutValue(TitleName2);
                            ws.Cells[dataIndex, splitLineIndex].Style.Font.Size = 10;
                            ws.Cells[dataIndex, splitLineIndex].Style.HorizontalAlignment = TextAlignmentType.Left;
                            ws.Cells[dataIndex, splitLineIndex].Style.VerticalAlignment = TextAlignmentType.Bottom;

                            dataIndex += 2;
                        }

                    }
                }

                //資料列上邊加上黑線
                //wb.Worksheets[0].Cells.CreateRange(index + 1, 0, 1, endIndex).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);

                //表格最右邊加上黑線
                //ws.Cells.CreateRange(index + 1, endIndex - 1, (dataIndex - index - 1), 1).SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

                index = dataIndex;

                //設定分頁
                ws.HPageBreaks.Add(index, endIndex);
            }

            if (dataIndex > 0)
            {
                //最後一頁的資料列下邊加上黑線
                ws.Cells.CreateRange(dataIndex - 1, 0, 1, endIndex).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            }
            else
                wb = new Workbook();

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
