using Aspose.Cells;
using FISCA.DSAUtil;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace Transfer_Excel
{
    internal class Report : IReport
    {
        #region IReport 成員

        private BackgroundWorker _BGWClassStudentDisciplineDetail;

        public void Print() //開始印出畫面
        {
            if (K12.Presentation.NLDPanels.Class.SelectedSource.Count == 0)
                return;

            DisciplineDetailForm form = new DisciplineDetailForm();
            if (form.ShowDialog() == DialogResult.OK)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化 班級學生獎懲明細(學生標題版)");

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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void _BGWClassStudentDisciplineDetail_DoWork(object sender, DoWorkEventArgs e)
        {
            string reportName = "班級學生獎懲明細(學生標題版)";

            object[] args = e.Argument as object[];

            DateTime startDate = (DateTime)args[0];
            DateTime endDate = (DateTime)args[1];
            int size = (int)args[2];
            bool IsInsertDate = (bool)args[3];
            Dictionary<string, bool> dic = (Dictionary<string, bool>)args[4];

            #region 快取資料

            List<ClassRecord> selectedClass = K12.Data.Class.SelectByIDs(K12.Presentation.NLDPanels.Class.SelectedSource);
            selectedClass = SortClassIndex.K12Data_ClassRecord(selectedClass);

            //學生ID查詢班級ID                                               
            Dictionary<string, string> studentClassDict = new Dictionary<string, string>();

            //學生ID查詢學生資訊
            Dictionary<string, StudentRecord> studentInfoDict = new Dictionary<string, StudentRecord>();

            //獎懲明細，Key為 ClassID -> StudentID -> OccurDate -> DisciplineType
            //Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, string>>>> allDisciplineDetail = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, string>>>>();

            //所有獎懲明細，按照班級編號, 及學生編號分開
            Dictionary<string, Dictionary<String, List<DisciplineEntity>>> allDisciplineRecords = new Dictionary<string, Dictionary<string, List<DisciplineEntity>>>();

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
                allDisciplineRecords.Add(aClass.ID, new Dictionary<string, List<DisciplineEntity>>());

                List<StudentRecord> classStudent = aClass.Students;

                foreach (StudentRecord student in classStudent)
                {
                    if (student.StatusStr == "一般" || student.StatusStr == "延修" || student.StatusStr == "輟學")
                    {
                        allStudentID.Add(student.ID);
                        studentClassDict.Add(student.ID, aClass.ID);
                        studentInfoDict.Add(student.ID, student);
                    }
                }
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
            helper.AddElement("Order", "RefStudentID", "asc");
            helper.AddElement("Order", "OccurDate", "asc");
            DSResponse dsrsp = QueryDiscipline.GetDiscipline(new DSRequest(helper));

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

                DisciplineEntity entity = new DisciplineEntity(var);
                string classID = studentClassDict[entity.StudentID];

                if (!allDisciplineRecords[classID].ContainsKey(entity.StudentID))
                    allDisciplineRecords[classID].Add(entity.StudentID, new List<DisciplineEntity>());

                allDisciplineRecords[classID][entity.StudentID].Add(new DisciplineEntity(var));

                totalNumber++;
            }

            #endregion

            #region 產生範本

            Workbook template = new Workbook(new MemoryStream(Properties.Resources.Sample),new LoadOptions(LoadFormat.Excel97To2003));

            Range tempStudent = template.Worksheets[0].Cells.CreateRange(0, 12, true);

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
            columnTable.Add("備註", colIndex++); //2020/2/18 - 新增

            //columnTable.Add("登錄日期", colIndex++);
            endIndex = colIndex;

            //prototype.Worksheets[0].Cells.CreateRange(0, 0, 1, endIndex).Merge();

            Range prototypeRow = prototype.Worksheets[0].Cells.CreateRange(1, 1, false);
            Range prototypeHeader = prototype.Worksheets[0].Cells.CreateRange(0, 1, false);

            #endregion

            #region  ==============  產生報表  ===================

            Workbook wb = new Workbook();
            wb.Copy(prototype);
            wb.CopyTheme(prototype);

            int currentProgress = 0; //目前列印進度
            //對於每一個班級
            foreach (ClassRecord classInfo in selectedClass)
            {
                //取得此班級的符合條件的所有獎懲紀錄 (每一筆紀錄是一個 DisciplineEntity 物件)
                Dictionary<String, List<DisciplineEntity>> classDisciplines = allDisciplineRecords[classInfo.ID];

                if (classDisciplines.Count <= 0)
                    continue;

                //每一個班級都建立一個 Worksheet
                int sheetindex = wb.Worksheets.AddCopy("Sheet1");
                Worksheet ws = wb.Worksheets[sheetindex];
                ws.Name = classInfo.Name + "班";
                //ws.Move(0);

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

                /* =============  開始列印此班級的獎懲紀錄   =============*/
                int rowIndex = 0;
                bool isFirstPage = true;

                //加入班級報表的標題 與日期區間
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

                wb.Worksheets[sheetindex].PageSetup.PrintTitleRows = "$1:$2";
                wb.Worksheets[sheetindex].PageSetup.SetFooter(1, "");

                //如果不是第一頁，就在上一頁的資料列下邊加黑線
                if (!isFirstPage)
                    ws.Cells.CreateRange(rowIndex - 1, 0, 1, endIndex).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                else
                    isFirstPage = !isFirstPage;

                //複製 Header
                tool.CopyStyle(ws.Cells.CreateRange(rowIndex, 2, false), prototypeHeader);

                ws.Cells[rowIndex, 0].PutValue(TitleName1);
                ws.Cells[rowIndex, 10].PutValue(TitleName2);

                //按照座號對學生編號清單做排序
                List<string> list = new List<string>();
                list.AddRange(classDisciplines.Keys);
                list.Sort(new SeatNoComparer(studentInfoDict));

                string[] fields = new string[] { "座號", "姓名", "日期", "大功", "小功", "嘉獎", "大過", "小過", "警告", "銷過", "銷過日期", "銷過事由", "事由" ,"備註"};

                rowIndex += 1;  //獎懲紀錄從第二列開始列印
                //對於每一位學生
                foreach (String studID in list)
                {
                    //下一位學生，應列印欄位名稱
                    tool.CopyStyle(ws.Cells.CreateRange(rowIndex, 1, false), prototypeRow);

                    int i = 0;
                    foreach (String field in fields)
                    {
                        ws.Cells[rowIndex, i].PutValue(fields[i]);
                        i++;
                    }
                    rowIndex += 1;

                    //列印獎懲資料
                    List<DisciplineEntity> studDisciplines = classDisciplines[studID];
                    foreach (DisciplineEntity de in studDisciplines)
                    {
                        tool.CopyStyle(ws.Cells.CreateRange(rowIndex, 1, false), prototypeRow);

                        StudentRecord studInfo = studentInfoDict[de.StudentID];
                        ws.Cells[rowIndex, 0].PutValue(studInfo.SeatNo);
                        ws.Cells[rowIndex, 1].PutValue(studInfo.Name);
                        ws.Cells[rowIndex, 2].PutValue(de.OccurDate.ToShortDateString());
                        if (de.MeritFlag == "1")
                        {
                            if (de.ACount != "0")
                                ws.Cells[rowIndex, 3].PutValue(de.ACount);
                            if (de.BCount != "0")
                                ws.Cells[rowIndex, 4].PutValue(de.BCount);
                            if (de.CCount != "0")
                                ws.Cells[rowIndex, 5].PutValue(de.CCount);
                            
                        }
                        else if (de.MeritFlag == "0")
                        {
                            if (de.ACount != "0")
                                ws.Cells[rowIndex, 6].PutValue(de.ACount);
                            if (de.BCount != "0")
                                ws.Cells[rowIndex, 7].PutValue(de.BCount);
                            if (de.CCount != "0")
                                ws.Cells[rowIndex, 8].PutValue(de.CCount);
                            ws.Cells[rowIndex, 9].PutValue(de.Clear);
                            ws.Cells[rowIndex, 10].PutValue(de.ClearDate);
                            ws.Cells[rowIndex, 11].PutValue(de.ClearReason);
                        }
                        ws.Cells[rowIndex, 12].PutValue(de.Reason);
                        ws.Cells[rowIndex, 13].PutValue(de.Remark);
                        rowIndex += 1;

                        currentProgress += 1;

                        _BGWClassStudentDisciplineDetail.ReportProgress((int)(((double)currentCount++ * 100.0) / (double)totalNumber));
                    }
                }
                //設定分頁
                //ws.HPageBreaks.Add(index, endIndex);

                //所有工作表資料文字置中
            }// end of each class
            #endregion
            //worksheet style Row自動調整、文字置中、換行
            //foreach (Worksheet sheet in wb.Worksheets)
            //{
            //foreach (Cell each in sheet.Cells)
            //{
            //    each.Style.HorizontalAlignment = TextAlignmentType.Center;
            //    each.Style.VerticalAlignment = TextAlignmentType.Center;
            //    each.Style.IsTextWrapped = true; //自動換行指令

            //}
            //sheet.AutoFitRows();
            //}

            //刪除多餘的worksheet
            wb.Worksheets.RemoveAt("Sheet1");
            wb.Worksheets.RemoveAt("Sheet2");
            wb.Worksheets.RemoveAt("Sheet3");

            string path = Path.Combine(Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".xls");
            e.Result = new object[] { reportName, path, wb };
        }
        #endregion
    }
}
