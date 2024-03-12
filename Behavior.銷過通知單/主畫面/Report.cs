using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using FISCA.DSAUtil;
using FISCA.Presentation.Controls;
using K12.Data;
using Framework.Feature;
using Aspose.Words;
using Behavior.Data;

namespace JHBehavior.銷過通知單
{
    internal class Report : IReport
    {

        private int DemeritAB; // 1 大過 等於 3 小過
        private int DemeritBC; // 1 小過 等於 3 警告
        private int MaxDemerit; //最小懲戒單位值
        private BackgroundWorker _BGWDisciplineNotification;

        private ConfigOBJ obj; //所有列印設定資訊

        private List<StudentRecord> SelectedStudents { get; set; }

        string entityName;

        public Report(string _entityName)
        {
            entityName = _entityName;
        }

        public void Print()
        {
            #region IReport 成員
            ClearDemeritDateRangeForm form = new ClearDemeritDateRangeForm();

            if (form.ShowDialog() == DialogResult.OK)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化銷過通知單...");

                #region 建立設定檔
                obj = new ConfigOBJ();
                obj.StartDate = form.StartDate;
                obj.EndDate = form.EndDate;
                obj.PrintHasRecordOnly = form.PrintHasRecordOnly;
                obj.Template = form.Template;
                obj.ReceiveName = form.ReceiveName;
                obj.ReceiveAddress = form.ReceiveAddress;
                obj.IsConditions = "銷過日期";
                obj.PrintStudentList = form.PrintStudentList;

                if (form.radioButton2.Checked)
                {
                    obj.IsConditions = "發生日期";
                }
                else if (form.radioButton3.Checked)
                {
                    obj.IsConditions = "登錄日期";
                }
                #endregion

                _BGWDisciplineNotification = new BackgroundWorker();
                _BGWDisciplineNotification.DoWork += new DoWorkEventHandler(_BGWDisciplineNotification_DoWork);
                _BGWDisciplineNotification.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CommonMethods.WordReport_RunWorkerCompleted);
                _BGWDisciplineNotification.ProgressChanged += new ProgressChangedEventHandler(CommonMethods.Report_ProgressChanged);
                _BGWDisciplineNotification.WorkerReportsProgress = true;
                _BGWDisciplineNotification.RunWorkerAsync();
            }
            #endregion
        }

        private void _BGWDisciplineNotification_DoWork(object sender, DoWorkEventArgs e)
        {
            #region Report
            if (entityName.ToLower() == "student") //學生模式
            {
                SelectedStudents = K12.Data.Student.SelectByIDs(K12.Presentation.NLDPanels.Student.SelectedSource);
            }
            else if (entityName.ToLower() == "class") //班級模式
            {
                SelectedStudents = new List<StudentRecord>();
                foreach (StudentRecord each in Student.SelectByClassIDs(K12.Presentation.NLDPanels.Class.SelectedSource))
                {
                    if (each.Status != StudentRecord.StudentStatus.一般)
                        continue;

                    SelectedStudents.Add(each);
                }
            }
            else
                throw new NotImplementedException();

            SelectedStudents.Sort(new Comparison<StudentRecord>(CommonMethods.ClassSeatNoComparer));
            #endregion

            string reportName = "銷過通知單";

            #region 快取資訊

            //超級資訊物件
            Dictionary<string, StudentOBJ> StudentSuperOBJ = new Dictionary<string, StudentOBJ>();
            //所有學生ID
            List<string> allStudentID = new List<string>();

            //學生人數
            int currentStudentCount = 1;
            int totalStudentNumber = 0;

            #endregion

            #region 依據 ClassID 建立班級學生清單

            //加總用
            Dictionary<string, int> StudMeritSum = new Dictionary<string, int>();

            foreach (StudentRecord aStudent in SelectedStudents)
            {
                if (!StudentSuperOBJ.ContainsKey(aStudent.ID))
                {
                    StudentSuperOBJ.Add(aStudent.ID, new StudentOBJ());
                }

                //學生ID清單
                if (!allStudentID.Contains(aStudent.ID))
                    allStudentID.Add(aStudent.ID);

                StudentSuperOBJ[aStudent.ID].student = aStudent;
                StudentSuperOBJ[aStudent.ID].TeacherName = aStudent.Class != null ? (aStudent.Class.Teacher != null ? aStudent.Class.Teacher.Name : "") : "";
                StudentSuperOBJ[aStudent.ID].ClassName = aStudent.Class != null ? aStudent.Class.Name : "";
                StudentSuperOBJ[aStudent.ID].SeatNo = aStudent.SeatNo.HasValue ? aStudent.SeatNo.Value.ToString() : "";
                StudentSuperOBJ[aStudent.ID].StudentNumber = aStudent.StudentNumber;

            }
            #endregion

            #region 取得獎懲資料(日期區間)
            List<DemeritRecord> DemeritList = new List<DemeritRecord>();

            if (obj.IsConditions == "銷過日期") //發生日期
            {
                DemeritList = Demerit.Select(allStudentID, null, null, null, null, obj.StartDate, obj.EndDate, null, null, null);
            }
            else if (obj.IsConditions == "發生日期")
            {
                DemeritList = Demerit.SelectByOccurDate(allStudentID, obj.StartDate, obj.EndDate);
            }
            else if (obj.IsConditions == "登錄日期")//登錄入期
            {
                DemeritList = Demerit.SelectByRegisterDate(allStudentID, obj.StartDate, obj.EndDate);
            }

            //依日期排序
            DemeritList.Sort(SortDateTime);

            foreach (DemeritRecord var in DemeritList)
            {
                string occurMonthDay = var.OccurDate.ToShortDateString();
                string reason = var.Reason;

                #region 懲戒
                if (var.Cleared == "是")
                {
                    ClearDemeritDetail cdd = new ClearDemeritDetail();


                    cdd.登錄日期 = occurMonthDay; //日期
                    cdd.懲戒事由 = reason; //事由
                    cdd.銷過事由 = var.ClearReason;
                    if (var.ClearDate.HasValue)
                    {
                        cdd.銷過日期 = var.ClearDate.Value.ToShortDateString();
                    }

                    StringBuilder detailString = new StringBuilder();
                    if (var.DemeritA != 0)
                    {
                        StudentSuperOBJ[var.RefStudentID].DemeritA += var.DemeritA.Value;
                        detailString.Append("大過：" + var.DemeritA.Value.ToString() + " ");
                    }
                    if (var.DemeritB != 0)
                    {
                        StudentSuperOBJ[var.RefStudentID].DemeritB += var.DemeritB.Value;
                        detailString.Append("小過：" + var.DemeritB.Value.ToString() + " ");
                    }
                    if (var.DemeritC != 0)
                    {
                        StudentSuperOBJ[var.RefStudentID].DemeritC += var.DemeritC.Value;
                        detailString.Append("警告：" + var.DemeritC.Value.ToString() + " ");
                    }
                    cdd.懲戒資料 = detailString.ToString();

                    StudentSuperOBJ[var.RefStudentID].DemeritStringList.Add(cdd);
                }
                #endregion

            }
            #endregion

            #region 取得學生通訊地址資料
            foreach (AddressRecord record in Address.SelectByStudentIDs(allStudentID))
            {
                if (obj.ReceiveAddress == "戶籍地址")
                {
                    if (!string.IsNullOrEmpty(record.PermanentAddress))
                        StudentSuperOBJ[record.RefStudentID].address = record.Permanent.County + record.Permanent.Town + record.Permanent.District + record.Permanent.Area + record.Permanent.Detail;

                    if (!string.IsNullOrEmpty(record.PermanentZipCode))
                    {
                        StudentSuperOBJ[record.RefStudentID].ZipCode = record.PermanentZipCode;

                        if (record.PermanentZipCode.Length >= 1)
                            StudentSuperOBJ[record.RefStudentID].ZipCode1 = record.PermanentZipCode.Substring(0, 1);
                        if (record.PermanentZipCode.Length >= 2)
                            StudentSuperOBJ[record.RefStudentID].ZipCode2 = record.PermanentZipCode.Substring(1, 1);
                        if (record.PermanentZipCode.Length >= 3)
                            StudentSuperOBJ[record.RefStudentID].ZipCode3 = record.PermanentZipCode.Substring(2, 1);
                        if (record.PermanentZipCode.Length >= 4)
                            StudentSuperOBJ[record.RefStudentID].ZipCode4 = record.PermanentZipCode.Substring(3, 1);
                        if (record.PermanentZipCode.Length >= 5)
                            StudentSuperOBJ[record.RefStudentID].ZipCode5 = record.PermanentZipCode.Substring(4, 1);
                    }

                }
                else if (obj.ReceiveAddress == "聯絡地址")
                {
                    if (!string.IsNullOrEmpty(record.MailingAddress))
                        StudentSuperOBJ[record.RefStudentID].address = record.Mailing.County + record.Mailing.Town + record.Mailing.District + record.Mailing.Area + record.Mailing.Detail; //再處理

                    if (!string.IsNullOrEmpty(record.MailingZipCode))
                    {
                        StudentSuperOBJ[record.RefStudentID].ZipCode = record.MailingZipCode;

                        if (record.MailingZipCode.Length >= 1)
                            StudentSuperOBJ[record.RefStudentID].ZipCode1 = record.MailingZipCode.Substring(0, 1);
                        if (record.MailingZipCode.Length >= 2)
                            StudentSuperOBJ[record.RefStudentID].ZipCode2 = record.MailingZipCode.Substring(1, 1);
                        if (record.MailingZipCode.Length >= 3)
                            StudentSuperOBJ[record.RefStudentID].ZipCode3 = record.MailingZipCode.Substring(2, 1);
                        if (record.MailingZipCode.Length >= 4)
                            StudentSuperOBJ[record.RefStudentID].ZipCode4 = record.MailingZipCode.Substring(3, 1);
                        if (record.MailingZipCode.Length >= 5)
                            StudentSuperOBJ[record.RefStudentID].ZipCode5 = record.MailingZipCode.Substring(4, 1);
                    }
                }
                else if (obj.ReceiveAddress == "其他地址")
                {
                    if (!string.IsNullOrEmpty(record.Address1Address))
                        StudentSuperOBJ[record.RefStudentID].address = record.Address1.County + record.Address1.Town + record.Address1.District + record.Address1.Area + record.Address1.Detail; //再處理

                    if (!string.IsNullOrEmpty(record.Address1ZipCode))
                    {
                        StudentSuperOBJ[record.RefStudentID].ZipCode = record.Address1ZipCode;

                        if (record.Address1ZipCode.Length >= 1)
                            StudentSuperOBJ[record.RefStudentID].ZipCode1 = record.Address1ZipCode.Substring(0, 1);
                        if (record.Address1ZipCode.Length >= 2)
                            StudentSuperOBJ[record.RefStudentID].ZipCode2 = record.Address1ZipCode.Substring(1, 1);
                        if (record.Address1ZipCode.Length >= 3)
                            StudentSuperOBJ[record.RefStudentID].ZipCode3 = record.Address1ZipCode.Substring(2, 1);
                        if (record.Address1ZipCode.Length >= 4)
                            StudentSuperOBJ[record.RefStudentID].ZipCode4 = record.Address1ZipCode.Substring(3, 1);
                        if (record.Address1ZipCode.Length >= 5)
                            StudentSuperOBJ[record.RefStudentID].ZipCode5 = record.Address1ZipCode.Substring(4, 1);
                    }
                }
            }
            #endregion

            #region 取得學生監護人父母親資料

            List<ParentRecord> ParentList = Parent.SelectByStudentIDs(allStudentID);

            foreach (ParentRecord record in ParentList)
            {
                StudentSuperOBJ[record.RefStudentID].CustodianName = record.CustodianName;
                StudentSuperOBJ[record.RefStudentID].FatherName = record.FatherName;
                StudentSuperOBJ[record.RefStudentID].MotherName = record.MotherName;
            }
            #endregion

            #region 產生報表

            Aspose.Words.Document template = new Aspose.Words.Document(obj.Template);
            template.MailMerge.Execute(
                new string[] { "學校名稱", "學校地址", "學校電話" },
                new object[] { School.ChineseName, School.Address, School.Telephone }
                );

            Aspose.Words.Document doc = new Aspose.Words.Document();
            doc.RemoveAllChildren();

            Aspose.Words.Node sectionNode = template.Sections[0].Clone();

            //取得學生人數
            totalStudentNumber = StudentSuperOBJ.Count;

            foreach (string student in StudentSuperOBJ.Keys)
            {
                //如果沒有學生就離開
                if (obj.PrintHasRecordOnly)
                {
                    if (StudentSuperOBJ.Count == 0)
                        continue;
                }

                //過濾不需要列印的學生
                if (StudentSuperOBJ[student].DemeritSum < MaxDemerit)
                    continue;

                if (StudentSuperOBJ[student].DemeritStringList.Count == 0)
                    continue;

                Aspose.Words.Document eachDoc = new Aspose.Words.Document();
                _run = new Run(eachDoc); //Cell寫入資料使用
                eachDoc.RemoveAllChildren();
                eachDoc.Sections.Add(eachDoc.ImportNode(sectionNode, true));

                //合併列印的資料
                Dictionary<string, object> mapping = new Dictionary<string, object>();

                StudentOBJ eachStudentInfo = StudentSuperOBJ[student];

                //學生資料
                mapping.Add("學生姓名", BehTool.SurrogatePairString(eachStudentInfo.student.Name));
                mapping.Add("班級", eachStudentInfo.ClassName);
                mapping.Add("座號", eachStudentInfo.SeatNo);
                mapping.Add("學號", eachStudentInfo.StudentNumber);
                mapping.Add("導師", eachStudentInfo.TeacherName);
                mapping.Add("資料期間", obj.StartDate.ToShortDateString() + " 至 " + obj.EndDate.ToShortDateString());

                //收件人資料
                if (obj.ReceiveName == "監護人姓名")
                    mapping.Add("收件人姓名", BehTool.SurrogatePairString(eachStudentInfo.CustodianName));
                else if (obj.ReceiveName == "父親姓名")
                    mapping.Add("收件人姓名", BehTool.SurrogatePairString(eachStudentInfo.FatherName));
                else if (obj.ReceiveName == "母親姓名")
                    mapping.Add("收件人姓名", BehTool.SurrogatePairString(eachStudentInfo.MotherName));
                else
                    mapping.Add("收件人姓名", BehTool.SurrogatePairString(eachStudentInfo.student.Name));

                //收件人地址資料
                mapping.Add("收件人地址", eachStudentInfo.address);
                mapping.Add("郵遞區號", eachStudentInfo.ZipCode);
                mapping.Add("0", eachStudentInfo.ZipCode1);
                mapping.Add("1", eachStudentInfo.ZipCode2);
                mapping.Add("2", eachStudentInfo.ZipCode3);
                mapping.Add("4", eachStudentInfo.ZipCode4);
                mapping.Add("5", eachStudentInfo.ZipCode5);

                mapping.Add("學年度", School.DefaultSchoolYear);
                mapping.Add("學期", School.DefaultSemester);

                //懲戒明細
                //object[] objectValues = new object[] { StudentSuperOBJ[student].DemeritStringList };
                //mapping.Add("銷過明細", objectValues);

                DocumentBuilder builder = new DocumentBuilder(eachDoc);
                builder.MoveToMergeField("銷過");
                Cell cell = (Cell)builder.CurrentParagraph.ParentNode;
                foreach (ClearDemeritDetail each in StudentSuperOBJ[student].DemeritStringList)
                {
                    if (Write(cell, each.登錄日期))
                    {
                        cell = cell.NextSibling as Cell;
                    }
                    if (Write(cell, each.懲戒資料))
                    {
                        cell = cell.NextSibling as Cell;
                    }
                    if (Write(cell, each.懲戒事由))
                    {
                        cell = cell.NextSibling as Cell;
                    }
                    if (Write(cell, each.銷過日期))
                    {
                        cell = cell.NextSibling as Cell;
                    }
                    if (Write(cell, each.銷過事由))
                    {
                        Row Nextrow = cell.ParentRow.NextSibling as Row; //取得下一個Row
                        if (Nextrow != null)
                        {
                            cell = Nextrow.FirstCell; //第一格Cell
                        }
                        else
                        {
                            break;
                        }
                    }
                }

                string[] keys = new string[mapping.Count];
                object[] values = new object[mapping.Count];
                int i = 0;
                foreach (string key in mapping.Keys)
                {
                    keys[i] = key;
                    values[i++] = mapping[key];
                }

                //合併列印
                eachDoc.MailMerge.MergeField += new Aspose.Words.Reporting.MergeFieldEventHandler(DisciplineNotification_MailMerge_MergeField);
                eachDoc.MailMerge.RemoveEmptyParagraphs = true;
                eachDoc.MailMerge.Execute(keys, values);

                Aspose.Words.Node eachSectionNode = eachDoc.Sections[0].Clone();
                doc.Sections.Add(doc.ImportNode(eachSectionNode, true));

                //回報進度
                _BGWDisciplineNotification.ReportProgress((int)(((double)currentStudentCount++ * 100.0) / (double)totalStudentNumber));
            }

            #endregion

            #region 產生學生清單

            Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook();
            if (obj.PrintStudentList)
            {
                if (obj.PrintStudentList)
                {

                    int CountRow = 0;
                    wb.Worksheets[0].Cells[CountRow, 0].PutValue("班級");
                    wb.Worksheets[0].Cells[CountRow, 1].PutValue("座號");
                    wb.Worksheets[0].Cells[CountRow, 2].PutValue("學號");
                    wb.Worksheets[0].Cells[CountRow, 3].PutValue("學生姓名");
                    wb.Worksheets[0].Cells[CountRow, 4].PutValue("收件人姓名");
                    wb.Worksheets[0].Cells[CountRow, 5].PutValue("地址");
                    CountRow++;
                    foreach (string each in StudentSuperOBJ.Keys)
                    {
                        if (StudentSuperOBJ[each].DemeritSum < MaxDemerit)
                            continue;

                        if (StudentSuperOBJ[each].DemeritStringList.Count == 0)
                            continue;

                        wb.Worksheets[0].Cells[CountRow, 0].PutValue(StudentSuperOBJ[each].ClassName);
                        wb.Worksheets[0].Cells[CountRow, 1].PutValue(StudentSuperOBJ[each].SeatNo);
                        wb.Worksheets[0].Cells[CountRow, 2].PutValue(StudentSuperOBJ[each].StudentNumber);
                        wb.Worksheets[0].Cells[CountRow, 3].PutValue(StudentSuperOBJ[each].student.Name);
                        //收件人資料
                        if (obj.ReceiveName == "監護人姓名")
                            wb.Worksheets[0].Cells[CountRow, 4].PutValue(StudentSuperOBJ[each].CustodianName);
                        else if (obj.ReceiveName == "父親姓名")
                            wb.Worksheets[0].Cells[CountRow, 4].PutValue(StudentSuperOBJ[each].FatherName);
                        else if (obj.ReceiveName == "母親姓名")
                            wb.Worksheets[0].Cells[CountRow, 4].PutValue(StudentSuperOBJ[each].MotherName);
                        else
                            wb.Worksheets[0].Cells[CountRow, 4].PutValue(StudentSuperOBJ[each].student.Name);

                        wb.Worksheets[0].Cells[CountRow, 5].PutValue(StudentSuperOBJ[each].ZipCode + " " + StudentSuperOBJ[each].address);
                        CountRow++;
                    }
                    wb.Worksheets[0].AutoFitColumns();
                }
            }
            #endregion

            string path = Path.Combine(Application.StartupPath, "Reports");
            string path2 = Path.Combine(Application.StartupPath, "Reports");

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");
            path2 = Path.Combine(path2, reportName + "(學生清單).xls");
            e.Result = new object[] { reportName, path, doc, path2, obj.PrintStudentList, wb };
        }

        private int SortDateTime(DemeritRecord x, DemeritRecord y)
        {
            return x.OccurDate.CompareTo(y.OccurDate);
        }

        private void DisciplineNotification_MailMerge_MergeField(object sender, Aspose.Words.Reporting.MergeFieldEventArgs e)
        {
            //#region MailMerge_MergeField
            //if (e.FieldName == "銷過明細")
            //{
            //    object[] objectValues = (object[])e.FieldValue;
            //    List<ClearDemeritDetail> eachStudentDisciplineDetail = (List<ClearDemeritDetail>)objectValues[0];

            //    Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(e.Document);

            //    builder.MoveToField(e.Field, false);
            //    builder.StartTable();
            //    builder.CellFormat.ClearFormatting();
            //    builder.CellFormat.Borders.ClearFormatting();
            //    builder.CellFormat.VerticalAlignment = Aspose.Words.CellVerticalAlignment.Center;
            //    builder.CellFormat.LeftPadding = 3.0;
            //    builder.RowFormat.LeftIndent = 0.0;
            //    builder.RowFormat.Height = 15.0;

            //    int rowNumber = 9;

            //    //if (eachStudentDisciplineDetail.Count > rowNumber * 2)
            //    //{
            //    //    rowNumber = eachStudentDisciplineDetail.Count / 2;
            //    //    rowNumber += eachStudentDisciplineDetail.Count % 2;
            //    //}

            //    //if (eachStudentDisciplineDetail.Count > rowNumber * 2)
            //    //{
            //    //    rowNumber += (eachStudentDisciplineDetail.Count - (rowNumber * 2)) / 2;
            //    //    rowNumber += (eachStudentDisciplineDetail.Count - (rowNumber * 2)) % 2;
            //    //}

            //    builder.CellFormat.Borders.Right.LineStyle = Aspose.Words.LineStyle.Single;
            //    builder.CellFormat.Borders.Right.Color = Color.Black;

            //    foreach (ClearDemeritDetail each in eachStudentDisciplineDetail)
            //    {
            //        Cell cell = (Cell)builder.CurrentParagraph.ParentNode;
            //        Write(cell, each.登錄日期);
            //        cell = cell.NextSibling as Cell;
            //        Write(cell, each.懲戒資料);
            //        cell = cell.NextSibling as Cell;
            //        Write(cell, each.懲戒事由);
            //        cell = cell.NextSibling as Cell;
            //        Write(cell, each.銷過日期);
            //        cell = cell.NextSibling as Cell;
            //        Write(cell, each.銷過事由);
            //    }



            //    builder.EndTable();

            //    e.Text = string.Empty;
            //}
            //#endregion
        }

        //移動使用
        private Run _run;

        /// <summary>
        /// 寫入資料
        /// </summary>
        private bool Write(Cell cell, string text)
        {
            if (cell == null)
                return false;

            if (text == null)
                return false;

            if (cell.FirstParagraph == null)
                cell.Paragraphs.Add(new Paragraph(cell.Document));
            cell.FirstParagraph.Runs.Clear();
            _run.Text = text;
            _run.Font.Size = 9;
            _run.Font.Name = "標楷體";
            cell.FirstParagraph.Runs.Add(_run.Clone(true));

            return true;
        }
    }
}
