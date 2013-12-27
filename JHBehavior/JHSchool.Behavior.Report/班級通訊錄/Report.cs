using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Aspose.Cells;
using FISCA.DSAUtil;
using Framework;
using JHSchool.Permrec;

namespace JHSchool.Behavior.Report.班級通訊錄
{
    internal class Report : IReport
    {
        #region IReport 成員

        private BackgroundWorker _BGWClassContactList;

        public void Print()
        {
            if (Class.Instance.SelectedList.Count == 0)
                return;

            FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化班級通訊錄...");

            _BGWClassContactList = new BackgroundWorker();
            _BGWClassContactList.DoWork += new DoWorkEventHandler(_BGWClassContactList_DoWork);
            _BGWClassContactList.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CommonMethods.ExcelReport_RunWorkerCompleted);
            _BGWClassContactList.ProgressChanged += new ProgressChangedEventHandler(CommonMethods.Report_ProgressChanged);
            _BGWClassContactList.WorkerReportsProgress = true;
            _BGWClassContactList.RunWorkerAsync();
        }

        void _BGWClassContactList_DoWork(object sender, DoWorkEventArgs e)
        {
            string reportName = "班級通訊錄";

            #region 快取所需要的資訊

            List<ClassRecord> selectedClasses = Class.Instance.SelectedList;

            selectedClasses = SortClassIndex.JHSchool_ClassRecord(selectedClasses);
            
            Dictionary<string, string> custodianName = new Dictionary<string, string>();
            Dictionary<string, string> mailingAddress = new Dictionary<string, string>();
            Dictionary<string, string> smsPhone = new Dictionary<string, string>();

            List<string> allClassId = new List<string>();

            int currentStudentNumber = 0;
            int allStudentNumber = 0;

            foreach (ClassRecord var in selectedClasses)
            {
                allStudentNumber += var.Students.Count;
                allClassId.Add(var.ID);
            }

            foreach (XmlElement element in (JHSchool.Compatibility.Feature.QueryStudent.GetDetailListByClassID(allClassId.ToArray()) as DSResponse).GetContent().GetElements("Student"))
            {
                XmlHelper xelement = new XmlHelper(element);
                string studentId = element.SelectSingleNode("@ID").InnerText;
                StringBuilder mailingAddressString = new StringBuilder();

                XmlElement addressElement = (XmlElement)element.SelectSingleNode("MailingAddress/AddressList/Address");
                if (addressElement != null)
                {
                    XmlHelper xaddress = new XmlHelper(addressElement);

                    mailingAddressString.Append(xaddress.GetString("ZipCode"));
                    mailingAddressString.Append(" ");
                    mailingAddressString.Append(xaddress.GetString("County"));
                    mailingAddressString.Append(xaddress.GetString("Town"));
                    mailingAddressString.Append(xaddress.GetString("DetailAddress"));

                    //mailingAddressString.Append(addressElement.SelectSingleNode("ZipCode").InnerText);
                    //mailingAddressString.Append(" ");
                    //mailingAddressString.Append(addressElement.SelectSingleNode("County").InnerText);
                    //mailingAddressString.Append(addressElement.SelectSingleNode("Town").InnerText);
                    //mailingAddressString.Append(addressElement.SelectSingleNode("DetailAddress").InnerText);
                }

                //custodianName.Add(studentId, element.SelectSingleNode("CustodianName").InnerText);
                custodianName.Add(studentId, xelement.GetString("CustodianName"));
                mailingAddress.Add(studentId, mailingAddressString.ToString());

                //smsPhone.Add(studentId, element.SelectSingleNode("SMSPhone").InnerText);
                smsPhone.Add(studentId, xelement.GetString("SMSPhone"));
            }

            #endregion

            #region 產生報表

            Workbook template = new Workbook();
            template.Open(new MemoryStream(ProjectResource.班級通訊錄), FileFormatType.Excel2003);

            Range tempRange = template.Worksheets[0].Cells.CreateRange(0, 32, false);

            Dictionary<string, int> sheets = new Dictionary<string, int>();
            Dictionary<string, int> sheetPageIndex = new Dictionary<string, int>();

            Workbook wb = new Workbook();
            wb.Open(new MemoryStream(ProjectResource.班級通訊錄), FileFormatType.Excel2003);

            Worksheet currentWorksheet;
            wb.Worksheets.Clear();

            int sheetIndex;
            int pageRow = 33;
            int pageCol = 8;
            int pageData = 30;

            SyncCacheData(selectedClasses);

            foreach (ClassRecord var in selectedClasses)
            {
                string gradeYear = var.GradeYear;
                List<StudentRecord> classStudent = var.Students.GetInSchoolStudents();
                //classStudent.Sort(new Comparison<StudentRecord>(CommonMethods.ClassSeatNoComparer));
                classStudent = SortClassIndex.JHSchool_StudentRecord(classStudent);
                if (!sheets.ContainsKey(gradeYear))
                {
                    int index;
                    int a;
                    //新增 sheet
                    index = wb.Worksheets.Add();
                    if (int.TryParse(gradeYear, out a))
                        wb.Worksheets[index].Name = gradeYear + " 年級";
                    else
                        wb.Worksheets[index].Name = gradeYear;
                    //sheet 列印設定
                    wb.Worksheets[index].PageSetup.Orientation = PageOrientationType.Landscape;
                    wb.Worksheets[index].PageSetup.TopMargin = 1.0;
                    wb.Worksheets[index].PageSetup.RightMargin = 0.6;
                    wb.Worksheets[index].PageSetup.BottomMargin = 1.0;
                    wb.Worksheets[index].PageSetup.LeftMargin = 0.6;
                    wb.Worksheets[index].PageSetup.CenterHorizontally = true;
                    wb.Worksheets[index].PageSetup.HeaderMargin = 0.8;
                    wb.Worksheets[index].PageSetup.FooterMargin = 0.8;
                    sheets.Add(gradeYear, index);
                    sheetPageIndex.Add(gradeYear, 0);

                    //複製 Template Column 寬度
                    for (int i = 0; i < pageCol; i++)
                    {
                        wb.Worksheets[index].Cells.CopyColumn(template.Worksheets[0].Cells, i, i);
                    }
                }

                //指定 sheet
                currentWorksheet = wb.Worksheets[sheets[gradeYear]];
                sheetIndex = sheetPageIndex[gradeYear];

                int currentPage = 1;
                int totalPage = (int)Math.Ceiling(((double)classStudent.Count / (double)pageData));

                int studentCount = 0;

                for (; currentPage <= totalPage; currentPage++)
                {
                    //複製 Template
                    currentWorksheet.Cells.CreateRange(sheetIndex, pageRow, false).Copy(tempRange);

                    //填寫班級基本資料
                    currentWorksheet.Cells[sheetIndex, 0].PutValue("班級：" + var.Name);
                    currentWorksheet.Cells[sheetIndex, 4].PutValue(School.ChineseName + " 學生通訊錄");
                    currentWorksheet.Cells[sheetIndex, 9].PutValue("班導師：" + (var.Teacher == null ? "" : var.Teacher.Name));

                    int dataIndex = sheetIndex + 2;

                    //填寫學生資料
                    for (int i = 0; i < pageData && studentCount < classStudent.Count; studentCount++, i++)
                    {
                        currentStudentNumber++;

                        PhoneRecord phone = classStudent[studentCount].GetPhone();

                        currentWorksheet.Cells[dataIndex, 0].PutValue(classStudent[studentCount].StudentNumber);
                        currentWorksheet.Cells[dataIndex, 1].PutValue(classStudent[studentCount].Class != null ? classStudent[studentCount].Class.Name : "");
                        currentWorksheet.Cells[dataIndex, 2].PutValue(classStudent[studentCount].SeatNo);
                        currentWorksheet.Cells[dataIndex, 3].PutValue(classStudent[studentCount].Name);
                        currentWorksheet.Cells[dataIndex, 4].PutValue(custodianName[classStudent[studentCount].ID]);
                        currentWorksheet.Cells[dataIndex, 5].PutValue(mailingAddress[classStudent[studentCount].ID]);
                        currentWorksheet.Cells[dataIndex, 6].PutValue(phone == null ? "" : phone.Permanent);
                        currentWorksheet.Cells[dataIndex, 7].PutValue(phone == null ? "" : phone.Contact);
                        currentWorksheet.Cells[dataIndex, 8].PutValue(smsPhone[classStudent[studentCount].ID]);
                        currentWorksheet.Cells[dataIndex, 9].PutValue(phone == null ? "" : phone.Phone1);
                        currentWorksheet.Cells[dataIndex, 10].PutValue(phone == null ? "" : phone.Phone2);
                        currentWorksheet.Cells[dataIndex, 11].PutValue(phone == null ? "" : phone.Phone3);

                        dataIndex++;

                        //回報進度
                        _BGWClassContactList.ReportProgress((int)((double)currentStudentNumber * 100.0 / (double)allStudentNumber));
                    }

                    //填寫頁數
                    currentWorksheet.Cells[sheetIndex + pageRow - 1, 9].PutValue("第 " + currentPage + " 頁 / 共 " + totalPage + " 頁");
                    //設定分頁
                    currentWorksheet.HPageBreaks.Add(sheetIndex + pageRow, pageCol);

                    //下一頁
                    sheetIndex += pageRow;
                }

                //把 sheet index 儲存起來
                sheetPageIndex[gradeYear] = sheetIndex;
            }
            wb.Worksheets.SortNames();

            #endregion

            string path = Path.Combine(Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".xlt");
            e.Result = new object[] { reportName, path, wb };
        }

        /// <summary>
        /// 取得選擇的班級裡所有學生電話資料。
        /// </summary>
        private static void SyncCacheData(List<ClassRecord> selectedClasses)
        {
            List<StudentRecord> synclist = new List<StudentRecord>();
            foreach (ClassRecord each in selectedClasses)
                synclist.AddRange(each.Students);
            synclist.SyncPhoneCache();
        }

        #endregion
    }
}
