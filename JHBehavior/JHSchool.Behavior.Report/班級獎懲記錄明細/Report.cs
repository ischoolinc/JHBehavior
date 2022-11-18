using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using Aspose.Cells;
using FISCA.DSAUtil;

namespace JHSchool.Behavior.Report.班級獎懲記錄明細
{
    internal class Report : IReport
    {
        #region IReport 成員

        private BackgroundWorker _BGWClassStudentDisciplineDetail;

        public void Print()
        {
            if (Class.Instance.SelectedList.Count == 0)
                return;

            DisciplineDetailForm form = new DisciplineDetailForm();
            if (form.ShowDialog() == DialogResult.OK)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化班級學生獎懲明細...");

                //開始日期,結束日期,列印尺寸,日期類型(發生日期or登錄日期)
                object[] args = new object[] { form.StartDate, form.EndDate, form.PaperSize, form.radioButton1.Checked, form.SetupDic };

                _BGWClassStudentDisciplineDetail = new BackgroundWorker();
                _BGWClassStudentDisciplineDetail.DoWork += new DoWorkEventHandler(_BGWClassStudentDisciplineDetail_DoWork);
                _BGWClassStudentDisciplineDetail.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CommonMethods.ExcelReport_RunWorkerCompleted);
                _BGWClassStudentDisciplineDetail.ProgressChanged += new ProgressChangedEventHandler(CommonMethods.Report_ProgressChanged);
                _BGWClassStudentDisciplineDetail.WorkerReportsProgress = true;
                _BGWClassStudentDisciplineDetail.RunWorkerAsync(args);
            }
        }

        #endregion


        #region 班級學生獎懲明細

        void _BGWClassStudentDisciplineDetail_DoWork(object sender, DoWorkEventArgs e)
        {
            string reportName = "班級獎懲記錄明細";

            object[] args = e.Argument as object[];

            DateTime startDate = (DateTime)args[0];
            DateTime endDate = (DateTime)args[1];
            int size = (int)args[2];
            bool IsInsertDate = (bool)args[3];
            Dictionary<string, bool> dic = (Dictionary<string, bool>)args[4];


            #region 快取資料

            List<ClassRecord> selectedClass = Class.Instance.SelectedList;
            selectedClass = SortClassIndex.JHSchool_ClassRecord(selectedClass);

            //學生ID查詢班級ID
            Dictionary<string, string> studentClassDict = new Dictionary<string, string>();

            //學生ID查詢學生資訊
            Dictionary<string, StudentRecord> studentInfoDict = new Dictionary<string, StudentRecord>();

            //獎懲明細，Key為 ClassID -> StudentID -> OccurDate -> DisciplineType
            Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, string>>>> allDisciplineDetail = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, string>>>>();

            //所有學生ID
            List<string> allStudentID = new List<string>();

            //獎懲筆數
            int currentCount = 1;
            int totalNumber = 0;

            //獎勵項目
            Dictionary<string, string> meritTable = new Dictionary<string, string>();
            meritTable.Add("大功", "A");
            meritTable.Add("小功", "B");
            meritTable.Add("嘉獎", "C");

            //懲戒項目
            Dictionary<string, string> demeritTable = new Dictionary<string, string>();
            demeritTable.Add("大過", "A");
            demeritTable.Add("小過", "B");
            demeritTable.Add("警告", "C");

            //紀錄每一個 Column 的 Index
            Dictionary<string, int> columnTable = new Dictionary<string, int>();

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

                allDisciplineDetail.Add(aClass.ID, new Dictionary<string, Dictionary<string, Dictionary<string, string>>>());
            }

            //取得獎懲資料 日期區間
            DSXmlHelper helper = new DSXmlHelper("Request");
            helper.AddElement("Field");
            helper.AddElement("Field", "All");
            helper.AddElement("Condition");
            foreach (string var in allStudentID)
            {
                helper.AddElement("Condition", "RefStudentID", var);
            }

            if (IsInsertDate) //登錄或是發生日期
            {
                helper.AddElement("Condition", "StartDate", startDate.ToShortDateString());
                helper.AddElement("Condition", "EndDate", endDate.ToShortDateString());
            }
            else
            {
                helper.AddElement("Condition", "StartRegisterDate", startDate.ToShortDateString());
                helper.AddElement("Condition", "EndRegisterDate", endDate.ToShortDateString());
            }

            helper.AddElement("Order");
            helper.AddElement("Order", "OccurDate", "asc");
            DSResponse dsrsp = JHSchool.Compatibility.Feature.Student.QueryDiscipline.GetDiscipline(new DSRequest(helper));

            foreach (XmlElement var in dsrsp.GetContent().GetElements("Discipline"))
            {
                #region 處理列印內容之判斷
                if (var.SelectSingleNode("MeritFlag").InnerText == "1") //獎勵
                {
                    if (!dic["獎勵"])
                        continue;
                }
                else if (var.SelectSingleNode("MeritFlag").InnerText == "0") //懲戒
                {
                    XmlElement demeritElement = (XmlElement)var.SelectSingleNode("Detail/Discipline/Demerit");

                    if (demeritElement.GetAttribute("Cleared") == "是")
                    {
                        if (!dic["懲戒銷過"])
                            continue;
                    }
                    else
                    {
                        if (!dic["懲戒未銷過"]) //未銷過也不要
                            continue;
                    }
                }
                else //其他
                {
                    continue;
                }

                #endregion

                string studentID = var.SelectSingleNode("RefStudentID").InnerText;
                DateTime occurDate = DateTime.Parse(var.SelectSingleNode("OccurDate").InnerText);
                string disciplineID = var.GetAttribute("ID");
                string occurDateID = occurDate.ToShortDateString() + "_" + disciplineID;
                string reason = var.SelectSingleNode("Reason").InnerText;
                string remark = var.SelectSingleNode("Remark").InnerText;
                string classID = studentClassDict[studentID];

                //string registerDate = var.SelectSingleNode("RegisterDate").InnerText;

                if (!allDisciplineDetail.ContainsKey(classID))
                    allDisciplineDetail.Add(classID, new Dictionary<string, Dictionary<string, Dictionary<string, string>>>());

                if (!allDisciplineDetail[classID].ContainsKey(studentID))
                    allDisciplineDetail[classID].Add(studentID, new Dictionary<string, Dictionary<string, string>>());

                if (!allDisciplineDetail[classID][studentID].ContainsKey(occurDateID))
                    allDisciplineDetail[classID][studentID].Add(occurDateID, new Dictionary<string, string>());

                //加入事由
                if (!allDisciplineDetail[classID][studentID][occurDateID].ContainsKey("事由"))
                    allDisciplineDetail[classID][studentID][occurDateID].Add("事由", reason);

                //加入備註 2020/1/8
                if (!allDisciplineDetail[classID][studentID][occurDateID].ContainsKey("備註"))
                    allDisciplineDetail[classID][studentID][occurDateID].Add("備註", remark);

                //if (!allDisciplineDetail[classID][studentID][occurDateID].ContainsKey("登錄日期"))
                //    allDisciplineDetail[classID][studentID][occurDateID].Add("登錄日期", registerDate);

                if (var.SelectSingleNode("MeritFlag").InnerText == "1") //判斷獎勵 or 懲戒
                {
                    #region 獎勵
                    XmlElement meritElement = (XmlElement)var.SelectSingleNode("Detail/Discipline/Merit");
                    if (meritElement == null) continue;

                    foreach (string merit in meritTable.Keys)
                    {
                        string times = meritElement.GetAttribute(meritTable[merit]);
                        if (times != "0")
                        {
                            if (!allDisciplineDetail[classID][studentID][occurDateID].ContainsKey(merit))
                                allDisciplineDetail[classID][studentID][occurDateID].Add(merit, "0");

                            allDisciplineDetail[classID][studentID][occurDateID][merit] = times;
                        }
                    }
                    #endregion
                }
                else
                {
                    #region 懲戒

                    XmlElement demeritElement = (XmlElement)var.SelectSingleNode("Detail/Discipline/Demerit");
                    if (demeritElement == null) continue;

                    string clearDate = "";
                    string clearReason = "";
                    if (demeritElement.GetAttribute("Cleared") == "是")
                    {
                        clearDate = demeritElement.GetAttribute("ClearDate");
                        clearReason = demeritElement.GetAttribute("ClearReason");
                        if (!allDisciplineDetail[classID][studentID][occurDateID].ContainsKey("銷過"))
                            allDisciplineDetail[classID][studentID][occurDateID].Add("銷過", "是");
                        if (!allDisciplineDetail[classID][studentID][occurDateID].ContainsKey("銷過日期"))
                            allDisciplineDetail[classID][studentID][occurDateID].Add("銷過日期", clearDate);
                        if (!allDisciplineDetail[classID][studentID][occurDateID].ContainsKey("銷過事由"))
                            allDisciplineDetail[classID][studentID][occurDateID].Add("銷過事由", clearReason);
                    }

                    foreach (string demerit in demeritTable.Keys)
                    {
                        string times = demeritElement.GetAttribute(demeritTable[demerit]);
                        if (times != "0")
                        {
                            if (!allDisciplineDetail[classID][studentID][occurDateID].ContainsKey(demerit))
                                allDisciplineDetail[classID][studentID][occurDateID].Add(demerit, "0");

                            allDisciplineDetail[classID][studentID][occurDateID][demerit] = times;
                        }
                    }
                    #endregion
                }

                totalNumber++;
            }

            #endregion

            #region 產生範本

            Workbook template = new Workbook(new MemoryStream(ProjectResource.班級獎懲記錄明細), new LoadOptions(LoadFormat.Excel97To2003));

            Range tempStudent = template.Worksheets[0].Cells.CreateRange(0, 12, false);

            Workbook prototype = new Workbook();
            prototype.Copy(template);
            prototype.CopyTheme(template);

            tool.CopyStyle(prototype.Worksheets[0].Cells.CreateRange(0, 12, true), tempStudent);

            int colIndex = 3;
            int endIndex = colIndex;
            foreach (string var in meritTable.Keys)
            {
                columnTable.Add(var, colIndex++);
            }
            foreach (string var in demeritTable.Keys)
            {
                columnTable.Add(var, colIndex++);
            }
            columnTable.Add("銷過", colIndex++);
            columnTable.Add("銷過日期", colIndex++);
            columnTable.Add("銷過事由", colIndex++);
            columnTable.Add("事由", colIndex++);
            columnTable.Add("備註", colIndex++);
            //columnTable.Add("登錄日期", colIndex++);
            endIndex = colIndex;

            //prototype.Worksheets[0].Cells.CreateRange(0, 0, 1, endIndex).Merge();

            Range prototypeRow = prototype.Worksheets[0].Cells.CreateRange(2, 1, false);
            Range prototypeHeader = prototype.Worksheets[0].Cells.CreateRange(0, 2, false);

            #endregion

            #region 產生報表

            Workbook wb = new Workbook();
            wb.Copy(prototype);
            wb.CopyTheme(prototype);

            foreach (ClassRecord classInfo in selectedClass)
            {

                //班級/學生/日期/詳細資料
                Dictionary<string, Dictionary<string, Dictionary<string, string>>> classDisciplineDetail = allDisciplineDetail[classInfo.ID];

                if (classDisciplineDetail.Count <= 0)
                    continue;

                int sheetindex = wb.Worksheets.AddCopy("Sheet1");
                Worksheet ws = wb.Worksheets[sheetindex];
                ws.Name = classInfo.Name + "班";

                #region 判斷紙張大小
                if (size == 0)
                {
                    ws.PageSetup.PaperSize = PaperSizeType.PaperA3;
                    ws.PageSetup.FitToPagesWide = 1;
                    ws.PageSetup.FitToPagesTall = 0;
                }
                else if (size == 1)
                {
                    ws.PageSetup.PaperSize = PaperSizeType.PaperA4;
                    ws.PageSetup.FitToPagesWide = 1;
                    ws.PageSetup.FitToPagesTall = 0;
                }
                else if (size == 2)
                {
                    ws.PageSetup.PaperSize = PaperSizeType.PaperB4;
                    ws.PageSetup.FitToPagesWide = 1;
                    ws.PageSetup.FitToPagesTall = 0;
                }
                #endregion

                int index = 0;
                int dataIndex = 0;

                string TitleName1 = classInfo.Name + "班　班級獎懲記錄明細";
                string TitleName2 = "";
                if (IsInsertDate)
                {
                    TitleName2 = "發生日期範圍：" + startDate.ToShortDateString() + " ~ " + endDate.ToShortDateString();
                }
                else
                {
                    TitleName2 = "登錄日期範圍：" + startDate.ToShortDateString() + " ~ " + endDate.ToShortDateString();
                }

                //如果不是第一頁，就在上一頁的資料列下邊加黑線
                if (index != 0)
                    ws.Cells.CreateRange(index - 1, 0, 1, endIndex).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

                //複製 Header
                tool.CopyStyle(ws.Cells.CreateRange(index, 2, true), prototypeHeader);

                dataIndex = index + 2;

                //填寫基本資料
                ws.Cells[index, 0].PutValue(TitleName1);
                //pageCount++;
                ws.Cells[index, 11].PutValue(TitleName2);
                //設定合併這個欄位
                //prototype.Worksheets[0].Cells.CreateRange(0, 11, 1, 2).Merge();

                wb.Worksheets[sheetindex].PageSetup.PrintTitleRows = "$1:$2";
                wb.Worksheets[sheetindex].PageSetup.SetFooter(1, "");

                List<string> list = new List<string>();
                list.AddRange(classDisciplineDetail.Keys);
                list.Sort(new SeatNoComparer(studentInfoDict));

                foreach (string studentID in list)
                {
                    //填寫資料
                    Dictionary<string, Dictionary<string, string>> studentDisciplineDetail = classDisciplineDetail[studentID];
                    foreach (string occurDateID in studentDisciplineDetail.Keys)
                    {
                        //計算筆數
                        //CountRows++;

                        //複製每一個 row
                        tool.CopyStyle(ws.Cells.CreateRange(dataIndex, 1, false), prototypeRow);

                        ws.Cells[dataIndex, 0].PutValue(studentInfoDict[studentID].SeatNo);
                        ws.Cells[dataIndex, 1].PutValue(studentInfoDict[studentID].Name);

                        string[] occurDateAndID = occurDateID.Split(new string[] { "_" }, StringSplitOptions.RemoveEmptyEntries);

                        ws.Cells[dataIndex, 2].PutValue(occurDateAndID[0]);

                        foreach (string type in studentDisciplineDetail[occurDateID].Keys)
                        {
                            if (columnTable.ContainsKey(type))
                                ws.Cells[dataIndex, columnTable[type]].PutValue(studentDisciplineDetail[occurDateID][type]);
                        }

                        dataIndex++;
                        _BGWClassStudentDisciplineDetail.ReportProgress((int)(((double)currentCount++ * 100.0) / (double)totalNumber));
                    }
                }

                index = dataIndex;

                //設定分頁
                ws.HorizontalPageBreaks.Add(index, endIndex);
            }


            #endregion
            wb.Worksheets.RemoveAt("Sheet1");
            wb.Worksheets.RemoveAt("Sheet2");
            wb.Worksheets.RemoveAt("Sheet3");

            string path = Path.Combine(Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".xlsx");
            e.Result = new object[] { reportName, path, wb };
        }
        #endregion
    }
}
