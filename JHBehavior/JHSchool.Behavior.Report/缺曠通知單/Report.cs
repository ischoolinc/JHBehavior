using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using Aspose.Words;
using FISCA.DSAUtil;
using K12.Data;
using K12.Data.Configuration;
using Framework;

namespace JHSchool.Behavior.Report.缺曠通知單
{
    internal class Report : IReport
    {
        public Report(string entityName)
        {
            if (entityName.ToLower() == "student")
            {
                SelectedStudents = Student.Instance.SelectedList;
            }
            else if (entityName.ToLower() == "class")
            {
                SelectedStudents = new List<StudentRecord>();
                foreach (ClassRecord each in Class.Instance.SelectedList)
                    SelectedStudents.AddRange(each.Students.GetInSchoolStudents());
            }
            else
                throw new NotImplementedException();

            SelectedStudents = SortClassIndex.JHSchool_StudentRecord(SelectedStudents);
        }

        #region IReport 成員

        public void Print()
        {
            if (SelectedStudents.Count <= 0) return;

            AbsenceNotificationSelectDateRangeForm form = new AbsenceNotificationSelectDateRangeForm();

            if (form.ShowDialog() == DialogResult.OK)
            {
                //讀取缺曠別 Preference
                Dictionary<string, List<string>> config = new Dictionary<string, List<string>>();

                //XmlElement preferenceData = CurrentUser.Instance.Preference["缺曠通知單_缺曠別設定"];
                Framework.ConfigData cd = User.Configuration["缺曠通知單_缺曠別設定"];
                XmlElement preferenceData = cd.GetXml("XmlData", null);

                if (preferenceData != null)
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

                FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化缺曠通知單...");

                object[] args = new object[] { form.StartDate, form.EndDate, form.PrintHasRecordOnly, form.Template, config, form.ReceiveName, form.ReceiveAddress, form.ConditionName, form.ConditionNumber, form.ConditionName2, form.ConditionNumber2, form.PrintStudentList };

                _BGWAbsenceNotification = new BackgroundWorker();
                _BGWAbsenceNotification.DoWork += new DoWorkEventHandler(_BGWAbsenceNotification_DoWork);
                _BGWAbsenceNotification.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CommonMethods.WordReport_RunWorkerCompleted);
                _BGWAbsenceNotification.ProgressChanged += new ProgressChangedEventHandler(CommonMethods.Report_ProgressChanged);
                _BGWAbsenceNotification.WorkerReportsProgress = true;
                _BGWAbsenceNotification.RunWorkerAsync(args);
            }
        }

        #endregion

        private BackgroundWorker _BGWAbsenceNotification;

        private List<StudentRecord> SelectedStudents { get; set; }

        private void _BGWAbsenceNotification_DoWork(object sender, DoWorkEventArgs e)
        {

            object[] args = e.Argument as object[];

            DateTime startDate = (DateTime)args[0];
            DateTime endDate = (DateTime)args[1];
            bool printHasRecordOnly = (bool)args[2];
            MemoryStream templateStream = (MemoryStream)args[3];
            Dictionary<string, List<string>> userDefinedConfig = (Dictionary<string, List<string>>)args[4];
            string receiveName = (string)args[5];
            string receiveAddress = (string)args[6];
            string condName = (string)args[7];
            int condNumber = int.Parse((string)args[8]);
            string condName2 = (string)args[9];
            int condNumber2 = int.Parse((string)args[10]);
            bool printStudentList = (bool)args[11];

            string reportName = "缺曠通知單" + startDate.ToString("yyyy/MM/dd") + "至" + endDate.ToString("yyyy/MM/dd");

            #region 快取資料

            //學生資訊
            Dictionary<string, Dictionary<string, string>> studentInfo = new Dictionary<string, Dictionary<string, string>>();

            //缺曠累計資料
            Dictionary<string, Dictionary<string, Dictionary<string, int>>> studentAbsence = new Dictionary<string, Dictionary<string, Dictionary<string, int>>>();

            //學期缺曠累計資料
            Dictionary<string, Dictionary<string, Dictionary<string, int>>> studentSemesterAbsence = new Dictionary<string, Dictionary<string, Dictionary<string, int>>>();

            //缺曠明細
            //Dictionary<string, List<string>> studentAbsenceDetail = new Dictionary<string, List<string>>();
            Dictionary<string, Dictionary<string, Dictionary<string, string>>> studentAbsenceDetail = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();

            //所有學生ID
            List<string> allStudentID = new List<string>();

            //學生人數
            int currentStudentCount = 1;
            int totalStudentNumber = 0;

            //Period List
            List<string> periodList = new List<string>();

            //Absence List
            Dictionary<string, string> absenceList = new Dictionary<string, string>();

            //使用者所選取的所有假別種類
            List<string> userDefinedAbsenceList = new List<string>();
            //if (condName == "")
            //{
            foreach (string kind in userDefinedConfig.Keys)
            {
                foreach (string type in userDefinedConfig[kind])
                {
                    if (!userDefinedAbsenceList.Contains(type))
                        userDefinedAbsenceList.Add(type);
                }
            }
            //}
            //else
            //{
            //    if (!userDefinedAbsenceList.Contains(condName))
            //        userDefinedAbsenceList.Add(condName);
            //}

            #region 取得所有學生ID
            foreach (StudentRecord aStudent in SelectedStudents)
            {
                string tname = string.Empty;
                if (aStudent.Class != null)
                    if (aStudent.Class.Teacher != null)
                        tname = aStudent.Class.Teacher.Name;

                //建立學生資訊，班級、座號、學號、姓名、導師
                string studentID = aStudent.ID;
                if (!studentInfo.ContainsKey(studentID))
                    studentInfo.Add(studentID, new Dictionary<string, string>());
                if (!studentInfo[studentID].ContainsKey("ClassName"))
                    studentInfo[studentID].Add("ClassName", aStudent.Class == null ? "" : aStudent.Class.Name);
                if (!studentInfo[studentID].ContainsKey("SeatNo"))
                    studentInfo[studentID].Add("SeatNo", aStudent.SeatNo);
                if (!studentInfo[studentID].ContainsKey("StudentNumber"))
                    studentInfo[studentID].Add("StudentNumber", aStudent.StudentNumber);
                if (!studentInfo[studentID].ContainsKey("Name"))
                    studentInfo[studentID].Add("Name", aStudent.Name);
                if (!studentInfo[studentID].ContainsKey("Teacher"))
                    studentInfo[studentID].Add("Teacher", tname);

                if (!allStudentID.Contains(studentID))
                    allStudentID.Add(studentID);

                totalStudentNumber++;
            }
            #endregion

            #region 取得 Period List

            List<K12.Data.PeriodMappingInfo> PeriodList = K12.Data.PeriodMapping.SelectAll();
            PeriodList.Sort(tool.SortPeriod);

            Dictionary<string, string> TestPeriodList = new Dictionary<string, string>();
            foreach (PeriodMappingInfo var in PeriodList)
            {
                if (!periodList.Contains(var.Name))
                    periodList.Add(var.Name);

                if (!TestPeriodList.ContainsKey(var.Name))
                {
                    TestPeriodList.Add(var.Name, var.Type);
                }
            }
            #endregion

            #region 取得 Absence List

            List<K12.Data.AbsenceMappingInfo> AbsenceList = K12.Data.AbsenceMapping.SelectAll();

            Dictionary<string, string> dirCC = new Dictionary<string, string>(); //代碼替換(新)
            foreach (K12.Data.AbsenceMappingInfo var in AbsenceList)
            {
                if (!absenceList.ContainsKey(var.Name))
                {
                    absenceList.Add(var.Name, var.Abbreviation);
                }

                if (!dirCC.ContainsKey(var.Name))
                {
                    dirCC.Add(var.Abbreviation, var.Name);
                }
            }
            #endregion

            #region 取得所有學生缺曠紀錄，日期區間
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
            DSResponse dsrsp = JHSchool.Compatibility.Feature.Student.QueryAttendance.GetAttendance(new DSRequest(helper));
            foreach (XmlElement var in dsrsp.GetContent().GetElements("Attendance"))
            {
                string studentID = var.SelectSingleNode("RefStudentID").InnerText;
                DateTime occurDate = DateTime.Parse(var.SelectSingleNode("OccurDate").InnerText);
                string occurDateString = occurDate.Year + "/" + occurDate.Month + "/" + occurDate.Day;

                //累計資料
                if (!studentAbsence.ContainsKey(studentID))
                    studentAbsence.Add(studentID, new Dictionary<string, Dictionary<string, int>>());

                //明細資料
                if (!studentAbsenceDetail.ContainsKey(studentID))
                    studentAbsenceDetail.Add(studentID, new Dictionary<string, Dictionary<string, string>>());
                if (!studentAbsenceDetail[studentID].ContainsKey(occurDateString))
                    studentAbsenceDetail[studentID].Add(occurDateString, new Dictionary<string, string>());

                foreach (XmlElement period in var.SelectNodes("Detail/Attendance/Period"))
                {
                    //<Period AbsenceType="事假" AttendanceType="一般">一</Period>

                    string absenceType = period.GetAttribute("AbsenceType");
                    string attendanceType = period.GetAttribute("AttendanceType"); //取得Period內容
                    string innerText = period.InnerText;

                    //如果attendanceType包含在userDefinedConfig內
                    if (!TestPeriodList.ContainsKey(innerText))
                        continue;

                    if (userDefinedConfig.ContainsKey(TestPeriodList[innerText]))
                    {
                        //absenceType是userDefinedConfig[attendanceType]裡面的原素
                        if (userDefinedConfig[TestPeriodList[innerText]].Contains(absenceType))
                        {
                            //如果studentAbsence[studentID]不包含attendanceType
                            if (!studentAbsence[studentID].ContainsKey(TestPeriodList[innerText]))
                            {
                                //在studentAbsence[studentID]加入一個attendanceType原素
                                studentAbsence[studentID].Add(TestPeriodList[innerText], new Dictionary<string, int>());
                            }

                            //如果studentAbsence[studentID][attendanceType]不包含absenceType
                            if (!studentAbsence[studentID][TestPeriodList[innerText]].ContainsKey(absenceType))
                            {
                                //在studentAbsence[studentID][attendanceType]加入一個absenceType,並且統計為0
                                studentAbsence[studentID][TestPeriodList[innerText]].Add(absenceType, 0);
                            }

                            //把統計值++
                            studentAbsence[studentID][TestPeriodList[innerText]][absenceType]++;
                        }
                    }

                    if (userDefinedAbsenceList.Contains(absenceType))
                    {
                        if (!studentAbsenceDetail[studentID][occurDateString].ContainsKey(innerText) && absenceList.ContainsKey(absenceType))
                            studentAbsenceDetail[studentID][occurDateString].Add(innerText, absenceList[absenceType]);
                    }
                }



                if (studentAbsenceDetail[studentID][occurDateString].Count <= 0) //如果統計小於0則移除該日內容
                    studentAbsenceDetail[studentID].Remove(occurDateString);
            }
            #endregion

            List<string> DelStudent = new List<string>(); //列印的學生


            #region 條件1
            if (condName != "") //如果不等於空就是要判斷啦
            {
                foreach (string each1 in studentAbsenceDetail.Keys) //取出一個學生
                {
                    int AbsenceCount = 0;
                    bool AbsenceBOOL = false;

                    foreach (string each2 in studentAbsenceDetail[each1].Keys) //取出一天
                    {
                        foreach (string each3 in studentAbsenceDetail[each1][each2].Keys) //取出一節內容
                        {
                            string each4 = studentAbsenceDetail[each1][each2][each3];
                            if (TestPeriodList.ContainsKey(each3))
                            {
                                if (userDefinedConfig.ContainsKey(TestPeriodList[each3]))
                                {
                                    if (dirCC[each4] == condName) //節次名稱為曠課
                                    {
                                        AbsenceCount++; //加1
                                    }

                                    if (AbsenceCount >= condNumber) //當曠課累積超過
                                    {
                                        AbsenceBOOL = true;
                                        if (!DelStudent.Contains(each1))
                                        {
                                            DelStudent.Add(each1); //把學生ID記下
                                        }
                                    }

                                    if (AbsenceBOOL)
                                        break;
                                }
                            }
                        }

                        if (AbsenceBOOL)
                            break;
                    }
                }
            }
            #endregion

            #region 條件2
            if (condName2 != "")
            {
                foreach (string each1 in studentAbsenceDetail.Keys) //取出一個學生
                {
                    int AbsenceCount = 0;
                    bool AbsenceBOOL = false;

                    foreach (string each2 in studentAbsenceDetail[each1].Keys) //取出一天
                    {
                        foreach (string each3 in studentAbsenceDetail[each1][each2].Keys) //取出一節內容
                        {
                            string each4 = studentAbsenceDetail[each1][each2][each3];

                            if (TestPeriodList.ContainsKey(each3))
                            {
                                if (userDefinedConfig.ContainsKey(TestPeriodList[each3]))
                                {
                                    if (dirCC[each4] == condName2) //節次名稱為曠課
                                    {
                                        AbsenceCount++; //加1
                                    }

                                    if (AbsenceCount >= condNumber2) //當曠課累積超過
                                    {
                                        AbsenceBOOL = true;
                                        if (!DelStudent.Contains(each1))
                                        {
                                            DelStudent.Add(each1); //把學生ID記下
                                        }
                                    }

                                    if (AbsenceBOOL)
                                        break;
                                }
                            }
                        }

                        if (AbsenceBOOL)
                            break;
                    }
                }
            }
            #endregion

            #region 無條件則全部列印
            if (condName == "" && condName2 == "")
            {
                foreach (string each1 in studentAbsenceDetail.Keys) //取出一個學生
                {
                    if (!DelStudent.Contains(each1))
                    {
                        DelStudent.Add(each1);
                    }
                }
            }
            #endregion


            #region 取得所有學生缺曠紀錄，學期累計
            helper = new DSXmlHelper("Request");
            helper.AddElement("Field");
            helper.AddElement("Field", "All");
            helper.AddElement("Condition");
            foreach (string var in allStudentID)
            {
                helper.AddElement("Condition", "RefStudentID", var);
            }
            helper.AddElement("Condition", "SchoolYear", School.DefaultSchoolYear.ToString());
            helper.AddElement("Condition", "Semester", School.DefaultSemester.ToString());
            helper.AddElement("Order");
            helper.AddElement("Order", "OccurDate", "asc");
            dsrsp = JHSchool.Compatibility.Feature.Student.QueryAttendance.GetAttendance(new DSRequest(helper));
            foreach (XmlElement var in dsrsp.GetContent().GetElements("Attendance"))
            {
                string studentID = var.SelectSingleNode("RefStudentID").InnerText;
                DateTime occurDate = DateTime.Parse(var.SelectSingleNode("OccurDate").InnerText);

                if (occurDate.CompareTo(endDate) == 1)
                    continue;

                string occurDateString = occurDate.Year + "/" + occurDate.Month + "/" + occurDate.Day;

                //累計資料
                if (!studentSemesterAbsence.ContainsKey(studentID))
                    studentSemesterAbsence.Add(studentID, new Dictionary<string, Dictionary<string, int>>());

                foreach (XmlElement period in var.SelectNodes("Detail/Attendance/Period"))
                {
                    string absenceType = period.GetAttribute("AbsenceType");
                    //string attendanceType = period.GetAttribute("AttendanceType");
                    string innerText = period.InnerText;

                    if (!TestPeriodList.ContainsKey(innerText))
                        continue;

                    if (userDefinedConfig.ContainsKey(TestPeriodList[innerText]))
                    {
                        if (userDefinedConfig[TestPeriodList[innerText]].Contains(absenceType))
                        {
                            if (!studentSemesterAbsence[studentID].ContainsKey(TestPeriodList[innerText]))
                                studentSemesterAbsence[studentID].Add(TestPeriodList[innerText], new Dictionary<string, int>());
                            if (!studentSemesterAbsence[studentID][TestPeriodList[innerText]].ContainsKey(absenceType))
                                studentSemesterAbsence[studentID][TestPeriodList[innerText]].Add(absenceType, 0);
                            studentSemesterAbsence[studentID][TestPeriodList[innerText]][absenceType]++;
                        }
                    }
                }
            }
            #endregion

            #region 取得學生通訊地址資料
            List<AddressRecord> AddressList = K12.Data.Address.SelectByStudentIDs(allStudentID);
            foreach (AddressRecord var in AddressList)
            {
                string studentID = var.RefStudentID;

                if (!studentInfo.ContainsKey(studentID))
                    studentInfo.Add(studentID, new Dictionary<string, string>());

                studentInfo[studentID].Add("Address", "");
                studentInfo[studentID].Add("ZipCode", "");
                studentInfo[studentID].Add("ZipCode1", "");
                studentInfo[studentID].Add("ZipCode2", "");
                studentInfo[studentID].Add("ZipCode3", "");
                studentInfo[studentID].Add("ZipCode4", "");
                studentInfo[studentID].Add("ZipCode5", "");

                if (receiveAddress == "戶籍地址")
                {
                    if (!string.IsNullOrEmpty(var.PermanentAddress))
                        studentInfo[studentID]["Address"] = var.PermanentCounty + var.PermanentTown + var.PermanentDistrict + var.PermanentArea + var.PermanentDetail;

                    if (!string.IsNullOrEmpty(var.PermanentZipCode))
                    {
                        studentInfo[studentID]["ZipCode"] = var.PermanentZipCode;

                        if (var.PermanentZipCode.Length >= 1)
                            studentInfo[studentID]["ZipCode1"] = var.PermanentZipCode.Substring(0, 1);
                        if (var.PermanentZipCode.Length >= 2)
                            studentInfo[studentID]["ZipCode2"] = var.PermanentZipCode.Substring(1, 1);
                        if (var.PermanentZipCode.Length >= 3)
                            studentInfo[studentID]["ZipCode3"] = var.PermanentZipCode.Substring(2, 1);
                        if (var.PermanentZipCode.Length >= 4)
                            studentInfo[studentID]["ZipCode4"] = var.PermanentZipCode.Substring(3, 1);
                        if (var.PermanentZipCode.Length >= 5)
                            studentInfo[studentID]["ZipCode5"] = var.PermanentZipCode.Substring(4, 1);
                    }

                }
                else if (receiveAddress == "聯絡地址")
                {
                    if (!string.IsNullOrEmpty(var.MailingAddress))
                        studentInfo[studentID]["Address"] = var.MailingCounty + var.MailingTown + var.MailingDistrict + var.MailingArea + var.MailingDetail;

                    if (!string.IsNullOrEmpty(var.MailingZipCode))
                    {
                        studentInfo[studentID]["ZipCode"] = var.MailingZipCode;

                        if (var.MailingZipCode.Length >= 1)
                            studentInfo[studentID]["ZipCode1"] = var.MailingZipCode.Substring(0, 1);
                        if (var.MailingZipCode.Length >= 2)
                            studentInfo[studentID]["ZipCode2"] = var.MailingZipCode.Substring(1, 1);
                        if (var.MailingZipCode.Length >= 3)
                            studentInfo[studentID]["ZipCode3"] = var.MailingZipCode.Substring(2, 1);
                        if (var.MailingZipCode.Length >= 4)
                            studentInfo[studentID]["ZipCode4"] = var.MailingZipCode.Substring(3, 1);
                        if (var.MailingZipCode.Length >= 5)
                            studentInfo[studentID]["ZipCode5"] = var.MailingZipCode.Substring(4, 1);
                    }
                }
                else if (receiveAddress == "其他地址")
                {
                    if (!string.IsNullOrEmpty(var.Address1Address))
                        studentInfo[studentID]["Address"] = var.Address1County + var.Address1Town + var.Address1District + var.Address1Area + var.Address1Detail;

                    if (!string.IsNullOrEmpty(var.Address1ZipCode))
                    {
                        studentInfo[studentID]["ZipCode"] = var.Address1ZipCode;

                        if (var.Address1ZipCode.Length >= 1)
                            studentInfo[studentID]["ZipCode1"] = var.Address1ZipCode.Substring(0, 1);
                        if (var.Address1ZipCode.Length >= 2)
                            studentInfo[studentID]["ZipCode2"] = var.Address1ZipCode.Substring(1, 1);
                        if (var.Address1ZipCode.Length >= 3)
                            studentInfo[studentID]["ZipCode3"] = var.Address1ZipCode.Substring(2, 1);
                        if (var.Address1ZipCode.Length >= 4)
                            studentInfo[studentID]["ZipCode4"] = var.Address1ZipCode.Substring(3, 1);
                        if (var.Address1ZipCode.Length >= 5)
                            studentInfo[studentID]["ZipCode5"] = var.Address1ZipCode.Substring(4, 1);
                    }
                }
            }
            #endregion

            #region 取得學生監護人父母親資料
            dsrsp = JHSchool.Compatibility.Feature.QueryStudent.GetMultiParentInfo(allStudentID.ToArray());
            foreach (XmlElement var in dsrsp.GetContent().GetElements("ParentInfo"))
            {
                string studentID = var.GetAttribute("StudentID");

                studentInfo[studentID].Add("CustodianName", var.SelectSingleNode("CustodianName").InnerText);
                studentInfo[studentID].Add("FatherName", var.SelectSingleNode("FatherName").InnerText);
                studentInfo[studentID].Add("MotherName", var.SelectSingleNode("MotherName").InnerText);
            }
            #endregion

            #endregion

            Document template = new Document(templateStream, "", LoadFormat.Doc, "");
            DocumentBuilder builder = new DocumentBuilder(template);

            //缺曠類別部份
            #region 缺曠類別部份
            builder.MoveToMergeField("缺曠類別");
            Table table = template.Sections[0].Body.Tables[0];
            Cell startCell = (Cell)builder.CurrentParagraph.ParentNode;
            Row startRow = (Row)startCell.ParentNode;

            double totalWidth = startCell.CellFormat.Width;
            int startRowIndex = table.IndexOf(startRow);
            int columnNumber = 0;

            foreach (List<string> var in userDefinedConfig.Values)
            {
                columnNumber += var.Count;
            }
            double columnWidth = totalWidth / columnNumber;

            for (int i = startRowIndex; i < startRowIndex + 4; i++)
            {
                table.Rows[i].RowFormat.HeightRule = HeightRule.Exactly;
                table.Rows[i].RowFormat.Height = 12;
            }

            foreach (string attendanceType in userDefinedConfig.Keys)
            {
                Cell newCell = new Cell(template);
                newCell.CellFormat.Width = userDefinedConfig[attendanceType].Count * columnWidth;
                newCell.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                newCell.CellFormat.WrapText = true;
                newCell.Paragraphs.Add(new Paragraph(template));
                newCell.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                newCell.Paragraphs[0].ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
                newCell.Paragraphs[0].ParagraphFormat.LineSpacing = 12;
                newCell.Paragraphs[0].Runs.Add(new Run(template, attendanceType));
                newCell.Paragraphs[0].Runs[0].Font.Size = 8;
                table.Rows[startRowIndex].Cells.Add(newCell.Clone(true));
                foreach (string absenceType in userDefinedConfig[attendanceType])
                {
                    newCell.CellFormat.Width = columnWidth;
                    newCell.Paragraphs[0].Runs[0].Text = absenceType;
                    table.Rows[startRowIndex + 1].Cells.Add(newCell.Clone(true));
                    newCell.Paragraphs[0].Runs[0].Text = "0";
                    table.Rows[startRowIndex + 2].Cells.Add(newCell.Clone(true));
                    table.Rows[startRowIndex + 3].Cells.Add(newCell.Clone(true));
                }
            }

            for (int i = startRowIndex; i < startRowIndex + 4; i++)
            {
                if (userDefinedConfig.Count > 0)
                    table.Rows[i].Cells[1].Remove();
                table.Rows[i].LastCell.CellFormat.Borders.Right.Color = Color.Black;
                table.Rows[i].LastCell.CellFormat.Borders.Right.LineWidth = 2.25;
            }
            #endregion


            #region 產生報表

            Document doc = new Document();
            doc.Sections.Clear();

            foreach (string studentID in studentInfo.Keys)
            {
                if (!DelStudent.Contains(studentID)) //如果不包含在內,就離開
                    continue;

                if (printHasRecordOnly)
                {
                    if (!studentAbsenceDetail.ContainsKey(studentID) || studentAbsenceDetail[studentID].Count == 0)
                    {
                        currentStudentCount++;
                        continue;
                    }
                }

                Document eachSection = new Document();
                eachSection.Sections.Clear();
                eachSection.Sections.Add(eachSection.ImportNode(template.Sections[0], true));

                //合併列印的資料
                Dictionary<string, object> mapping = new Dictionary<string, object>();
                Dictionary<string, string> eachStudentInfo = studentInfo[studentID];

                //學校資訊
                mapping.Add("學校名稱", School.ChineseName);
                mapping.Add("學校地址", School.Address);
                mapping.Add("學校電話", School.Telephone);

                //學生資料
                mapping.Add("系統編號", "系統編號{" + studentID + "}");
                mapping.Add("學生姓名", eachStudentInfo["Name"]);
                mapping.Add("班級", eachStudentInfo["ClassName"]);
                mapping.Add("座號", eachStudentInfo["SeatNo"]);
                mapping.Add("學號", eachStudentInfo["StudentNumber"]);
                mapping.Add("導師", eachStudentInfo["Teacher"]);
                mapping.Add("資料期間", startDate.ToShortDateString() + " 至 " + endDate.ToShortDateString());

                //收件人資料
                if (receiveName == "監護人姓名")
                    mapping.Add("收件人姓名", eachStudentInfo["CustodianName"]);
                else if (receiveName == "父親姓名")
                    mapping.Add("收件人姓名", eachStudentInfo["FatherName"]);
                else if (receiveName == "母親姓名")
                    mapping.Add("收件人姓名", eachStudentInfo["MotherName"]);
                else
                    mapping.Add("收件人姓名", eachStudentInfo["Name"]);

                //收件人地址資料
                mapping.Add("收件人地址", eachStudentInfo["Address"]);
                mapping.Add("郵遞區號", eachStudentInfo["ZipCode"]);
                mapping.Add("0", eachStudentInfo["ZipCode1"]);
                mapping.Add("1", eachStudentInfo["ZipCode2"]);
                mapping.Add("2", eachStudentInfo["ZipCode3"]);
                mapping.Add("4", eachStudentInfo["ZipCode4"]);
                mapping.Add("5", eachStudentInfo["ZipCode5"]);

                mapping.Add("學年度", School.DefaultSchoolYear);
                mapping.Add("學期", School.DefaultSemester);

                //缺曠明細
                if (studentAbsenceDetail.ContainsKey(studentID))
                {
                    object[] objectValues = new object[] { studentAbsenceDetail[studentID], periodList };
                    mapping.Add("缺曠明細", objectValues);

                }
                else
                    mapping.Add("缺曠明細", null);

                string[] keys = new string[mapping.Count];
                object[] values = new object[mapping.Count];
                int i = 0;
                foreach (string key in mapping.Keys)
                {
                    keys[i] = key;
                    values[i++] = mapping[key];
                }

                //合併列印
                eachSection.MailMerge.MergeField += new Aspose.Words.Reporting.MergeFieldEventHandler(AbsenceNotification_MailMerge_MergeField);
                eachSection.MailMerge.RemoveEmptyParagraphs = true;
                eachSection.MailMerge.Execute(keys, values);

                //填寫缺曠記錄
                Table eachTable = eachSection.Sections[0].Body.Tables[0];
                int columnIndex = 1;
                foreach (string attendanceType in userDefinedConfig.Keys)
                {
                    foreach (string absenceType in userDefinedConfig[attendanceType])
                    {
                        string dataValue = "0";
                        string semesterDataValue = "0";
                        if (studentAbsence.ContainsKey(studentID) && studentAbsence[studentID].ContainsKey(attendanceType))
                        {
                            if (studentAbsence[studentID][attendanceType].ContainsKey(absenceType))
                                dataValue = studentAbsence[studentID][attendanceType][absenceType].ToString();
                        }
                        if (studentSemesterAbsence.ContainsKey(studentID) && studentSemesterAbsence[studentID].ContainsKey(attendanceType))
                        {
                            if (studentSemesterAbsence[studentID][attendanceType].ContainsKey(absenceType))
                                semesterDataValue = studentSemesterAbsence[studentID][attendanceType][absenceType].ToString();
                        }
                        eachTable.Rows[startRowIndex + 3].Cells[columnIndex].Paragraphs[0].Runs[0].Text = dataValue;
                        eachTable.Rows[startRowIndex + 2].Cells[columnIndex].Paragraphs[0].Runs[0].Text = semesterDataValue;
                        columnIndex++;
                    }
                }

                doc.Sections.Add(doc.ImportNode(eachSection.Sections[0], true));

                //回報進度
                _BGWAbsenceNotification.ReportProgress((int)(((double)currentStudentCount++ * 100.0) / (double)totalStudentNumber));
            }

            #endregion

            #region 產生學生清單
            Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook();
            if (printStudentList)
            {
                int CountRow = 0;
                wb.Worksheets[0].Cells[CountRow, 0].PutValue("班級");
                wb.Worksheets[0].Cells[CountRow, 1].PutValue("座號");
                wb.Worksheets[0].Cells[CountRow, 2].PutValue("學號");
                wb.Worksheets[0].Cells[CountRow, 3].PutValue("學生姓名");
                wb.Worksheets[0].Cells[CountRow, 4].PutValue("收件人姓名");
                wb.Worksheets[0].Cells[CountRow, 5].PutValue("地址");
                CountRow++;
                foreach (string each in studentInfo.Keys)
                {
                    if (!DelStudent.Contains(each)) //如果不包含在內,就離開
                        continue;
                    if (!studentAbsenceDetail.ContainsKey(each) || studentAbsenceDetail[each].Count == 0)
                        continue;

                    wb.Worksheets[0].Cells[CountRow, 0].PutValue(studentInfo[each]["ClassName"]);
                    wb.Worksheets[0].Cells[CountRow, 1].PutValue(studentInfo[each]["SeatNo"]);
                    wb.Worksheets[0].Cells[CountRow, 2].PutValue(studentInfo[each]["StudentNumber"]);
                    wb.Worksheets[0].Cells[CountRow, 3].PutValue(studentInfo[each]["Name"]);
                    //收件人資料
                    if (receiveName == "監護人姓名")
                        wb.Worksheets[0].Cells[CountRow, 4].PutValue(studentInfo[each]["CustodianName"]);
                    else if (receiveName == "父親姓名")
                        wb.Worksheets[0].Cells[CountRow, 4].PutValue(studentInfo[each]["FatherName"]);
                    else if (receiveName == "母親姓名")
                        wb.Worksheets[0].Cells[CountRow, 4].PutValue(studentInfo[each]["MotherName"]);
                    else
                        wb.Worksheets[0].Cells[CountRow, 4].PutValue(studentInfo[each]["Name"]);

                    wb.Worksheets[0].Cells[CountRow, 5].PutValue(studentInfo[each]["ZipCode"] + " " + studentInfo[each]["Address"]);
                    CountRow++;
                }
                wb.Worksheets[0].AutoFitColumns();
            }
            #endregion

            string path = Path.Combine(Application.StartupPath, "Reports");
            string path2 = Path.Combine(Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");
            path2 = Path.Combine(path2, reportName + "(學生清單).xls");
            e.Result = new object[] { reportName, path, doc, path2, printStudentList, wb };
        }

        private void AbsenceNotification_MailMerge_MergeField(object sender, Aspose.Words.Reporting.MergeFieldEventArgs e)
        {
            #region 缺曠明細
            if (e.FieldName == "缺曠明細")
            {
                if (e.FieldValue == null)
                    return;

                object[] objectValues = (object[])e.FieldValue;
                Dictionary<string, Dictionary<string, string>> studentAbsenceDetail = (Dictionary<string, Dictionary<string, string>>)objectValues[0];
                List<string> periodList = (List<string>)objectValues[1];

                DocumentBuilder builder = new DocumentBuilder(e.Document);

                #region 缺曠明細部份
                builder.MoveToField(e.Field, false);
                Cell detailStartCell = (Cell)builder.CurrentParagraph.ParentNode;
                Row detailStartRow = (Row)detailStartCell.ParentNode;
                int detailStartRowIndex = e.Document.Sections[0].Body.Tables[0].IndexOf(detailStartRow);

                Table detailTable = builder.StartTable();
                builder.CellFormat.Borders.Left.LineWidth = 0.5;
                builder.CellFormat.Borders.Right.LineWidth = 0.5;

                builder.RowFormat.HeightRule = HeightRule.Auto;
                builder.RowFormat.Height = 12;
                builder.RowFormat.Alignment = RowAlignment.Center;

                int rowNumber = 4; //共4個Row,依缺曠天數進行調整
                if (studentAbsenceDetail.Count > rowNumber * 3)
                {
                    rowNumber = studentAbsenceDetail.Count / 3;
                    if (studentAbsenceDetail.Count % 3 > 0)
                        rowNumber++;
                }

                #region 暫解阿!!
                int TestPeriodListCount = periodList.Count;
                if (periodList.Count < 10)
                {
                    TestPeriodListCount = 10;
                }
                else
                {
                    TestPeriodListCount = periodList.Count;
                }
                #endregion

                builder.InsertCell();

                #region 填入日期 & 節次
                for (int i = 0; i < 3; i++)
                {
                    builder.CellFormat.Borders.Right.Color = Color.Black;
                    builder.CellFormat.Borders.Left.Color = Color.Black;
                    builder.CellFormat.Width = 20;
                    builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                    builder.CellFormat.Borders.LineStyle = LineStyle.Single;
                    builder.Write("日期");
                    builder.InsertCell();

                    for (int j = 0; j < TestPeriodListCount; j++)
                    {
                        builder.CellFormat.Borders.Right.Color = Color.Black;
                        builder.CellFormat.Borders.Left.Color = Color.Black;
                        builder.CellFormat.Borders.LineStyle = LineStyle.Dot;
                        builder.CellFormat.Width = 9;
                        builder.CellFormat.WrapText = true;
                        builder.CellFormat.LeftPadding = 0.5;
                        if (j < periodList.Count)
                        {
                            builder.Write(periodList[j]); //寫入節次名稱
                        }
                        builder.InsertCell();
                    }
                }
                #endregion

                builder.EndRow();

                #region 建立每日格數
                for (int x = 0; x < rowNumber; x++)
                {
                    builder.CellFormat.Borders.Right.Color = Color.Black;
                    builder.CellFormat.Borders.Left.Color = Color.Black;
                    builder.CellFormat.Borders.Left.LineWidth = 0.5;
                    builder.CellFormat.Borders.Right.LineWidth = 0.5;
                    builder.CellFormat.Borders.Top.LineWidth = 0.5;
                    builder.CellFormat.Borders.Bottom.LineWidth = 0.5;
                    builder.CellFormat.Borders.LineStyle = LineStyle.Dot;
                    builder.RowFormat.HeightRule = HeightRule.Exactly;
                    builder.RowFormat.Height = 12;
                    builder.RowFormat.Alignment = RowAlignment.Center;
                    builder.InsertCell();

                    for (int i = 0; i < 3; i++)
                    {
                        builder.CellFormat.Borders.Left.LineStyle = LineStyle.Single;
                        builder.CellFormat.Width = 20;
                        builder.Write("");
                        builder.InsertCell();

                        builder.CellFormat.Borders.LineStyle = LineStyle.Dot;

                        for (int j = 0; j < TestPeriodListCount; j++)
                        {
                            builder.CellFormat.Width = 9;
                            builder.Write("");
                            builder.InsertCell();
                        }
                    }

                    builder.EndRow();
                }
                #endregion

                builder.EndTable();

                foreach (Cell var in detailTable.Rows[0].Cells)
                {
                    var.Paragraphs[0].ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
                    var.Paragraphs[0].ParagraphFormat.LineSpacing = 9;
                }
                #endregion

                #region 填寫缺曠明細
                int eachDetailRowIndex = 0;
                int eachDetailColIndex = 0;

                foreach (string date in studentAbsenceDetail.Keys)
                {
                    int eachDetailPeriodColIndex = eachDetailColIndex + 1;
                    string[] splitDate = date.Split('/');
                    Paragraph dateParagraph = detailTable.Rows[eachDetailRowIndex + 1].Cells[eachDetailColIndex].Paragraphs[0];
                    dateParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    dateParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
                    dateParagraph.ParagraphFormat.LineSpacing = 9;

                    dateParagraph.Runs.Clear();
                    dateParagraph.Runs.Add(new Run(e.Document));
                    dateParagraph.Runs[0].Font.Size = 8;
                    dateParagraph.Runs[0].Text = splitDate[1] + "/" + splitDate[2];

                    foreach (string period in periodList)
                    {
                        string dataValue = "";
                        if (studentAbsenceDetail[date].ContainsKey(period))
                        {
                            dataValue = studentAbsenceDetail[date][period];
                            Cell miniCell = detailTable.Rows[eachDetailRowIndex + 1].Cells[eachDetailPeriodColIndex];
                            miniCell.Paragraphs.Clear();
                            miniCell.Paragraphs.Add(dateParagraph.Clone(true));
                            miniCell.Paragraphs[0].Runs[0].Font.Size = 14 - (int)(TestPeriodListCount / 2); //依表格多寡縮小文字
                            miniCell.Paragraphs[0].Runs[0].Text = dataValue;
                        }
                        eachDetailPeriodColIndex++;
                    }
                    eachDetailRowIndex++;
                    if (eachDetailRowIndex >= rowNumber)
                    {
                        eachDetailRowIndex = 0;
                        eachDetailColIndex += (TestPeriodListCount + 1);
                    }
                }
                #endregion

                e.Text = string.Empty;
            }
            #endregion
        }
    }
}
