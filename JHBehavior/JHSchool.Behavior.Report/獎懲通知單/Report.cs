using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Behavior.Data;
using FISCA.DSAUtil;
using Framework;
using Framework.Feature;
using K12.Data;

namespace JHSchool.Behavior.Report.獎勵懲戒通知單
{
    internal class Report : IReport
    {
        private int DemeritAB; // 1 大過 等於 3 小過
        private int DemeritBC; // 1 小過 等於 3 警告
        private int MaxDemerit; //最小懲戒單位值

        private int MeritAB;
        private int MeritBC;
        private int MaxMerit;

        string MeritDemerit = "";


        private BackgroundWorker _BGWDisciplineNotification;

        private List<StudentRecord> SelectedStudents { get; set; }

        public Report(string entityName)
        {
            #region Report
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
            #endregion
        }

        public void Print()
        {
            #region IReport 成員
            DisciplineNotificationSelectDateRangeForm form = new DisciplineNotificationSelectDateRangeForm();

            if (form.ShowDialog() == DialogResult.OK)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化獎懲通知單...");

                object[] args = new object[] { form.StartDate, form.EndDate, form.PrintHasRecordOnly,
                    form.Template, form.ReceiveName, form.ReceiveAddress,
                    form.ConditionName, form.ConditionNumber, form.radioButton1.Checked,
                    form.PrintStudentList , form.PrintRemark };

                _BGWDisciplineNotification = new BackgroundWorker();
                _BGWDisciplineNotification.DoWork += new DoWorkEventHandler(_BGWDisciplineNotification_DoWork);
                _BGWDisciplineNotification.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CommonMethods.WordReport_RunWorkerCompleted);
                _BGWDisciplineNotification.ProgressChanged += new ProgressChangedEventHandler(CommonMethods.Report_ProgressChanged);
                _BGWDisciplineNotification.WorkerReportsProgress = true;
                _BGWDisciplineNotification.RunWorkerAsync(args);
            }
            #endregion
        }

        private void GetReduceList()
        {
            #region 取得獎懲對照表
            DSResponse dsrsp = Config.GetMDReduce();
            if (!dsrsp.HasContent)
            {
                MsgBox.Show("取得對照表失敗 : " + dsrsp.GetFault().Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            DSXmlHelper helper = dsrsp.GetContent();

            string jb;

            jb = helper.GetText("Merit/AB");
            if (!int.TryParse(jb, out MeritAB))
            {
                MsgBox.Show("獎懲對照表有誤");
                return;
            }

            jb = helper.GetText("Merit/BC");
            if (!int.TryParse(jb, out MeritBC))
            {
                MsgBox.Show("獎懲對照表有誤");
                return;
            }


            jb = helper.GetText("Merit/AB");
            if (!int.TryParse(jb, out DemeritAB))
            {
                MsgBox.Show("獎懲對照表有誤");
                return;
            }

            jb = helper.GetText("Demerit/BC");
            if (!int.TryParse(jb, out DemeritBC))
            {
                MsgBox.Show("獎懲對照表有誤");
                return;
            }

            #endregion
        }

        private void ChengeDemerit(string CondName, int CondNumber)
        {
            #region 取得最小單位值

            MaxDemerit = 0; //最小單位值

            if (CondName == "大過")
            {
                MeritDemerit = "懲戒";
                MaxDemerit = CondNumber * DemeritAB;
                MaxDemerit = MaxDemerit * DemeritBC;
            }
            else if (CondName == "小過")
            {
                MeritDemerit = "懲戒";
                MaxDemerit = CondNumber * DemeritBC;
            }
            else if (CondName == "警告")
            {
                MeritDemerit = "懲戒";
                MaxDemerit = CondNumber;
            }
            else if (CondName == "大功")
            {
                MeritDemerit = "獎勵";
                MaxMerit = CondNumber * MeritAB;
                MaxMerit = MaxMerit * MeritBC;
            }
            else if (CondName == "小功")
            {
                MeritDemerit = "獎勵";
                MaxMerit = CondNumber * MeritBC;
            }
            else if (CondName == "嘉獎")
            {
                MeritDemerit = "獎勵";
                MaxMerit = CondNumber;
            }
            #endregion
        }

        private int GetDemeritType(XmlElement xml)
        {
            #region 將懲戒換算為最小單位值
            string AA = xml.GetAttribute("A");
            string BB = xml.GetAttribute("B");
            string CC = xml.GetAttribute("C");

            int SumNum1 = 0;
            int SumNum2 = 0;
            int SumNum3 = 0;

            if (int.Parse(AA) != 0)
            {
                SumNum1 = int.Parse(AA) * DemeritAB; //轉換成小功
                SumNum1 = SumNum1 * DemeritBC; //轉換成警告
            }

            if (int.Parse(BB) != 0)
            {
                SumNum2 = int.Parse(BB) * DemeritBC; //轉換成警告
            }

            if (int.Parse(CC) != 0)
            {
                SumNum3 = int.Parse(CC); //轉換成警告
            }


            return SumNum1 + SumNum2 + SumNum3;
            #endregion
        }

        private int GetMeritType(XmlElement xml)
        {
            #region 將獎勵換算為最小單位值
            string AA = xml.GetAttribute("A");
            string BB = xml.GetAttribute("B");
            string CC = xml.GetAttribute("C");

            int SumNum1 = 0;
            int SumNum2 = 0;
            int SumNum3 = 0;

            if (int.Parse(AA) != 0)
            {
                SumNum1 = int.Parse(AA) * MeritAB; //轉換成小功
                SumNum1 = SumNum1 * MeritBC; //轉換成警告
            }

            if (int.Parse(BB) != 0)
            {
                SumNum2 = int.Parse(BB) * MeritBC; //轉換成警告
            }

            if (int.Parse(CC) != 0)
            {
                SumNum3 = int.Parse(CC); //轉換成警告
            }


            return SumNum1 + SumNum2 + SumNum3;
            #endregion
        }

        private void _BGWDisciplineNotification_DoWork(object sender, DoWorkEventArgs e)
        {
            #region 表頭
            GetReduceList(); //獎懲對照表

            object[] args = e.Argument as object[];

            DateTime startDate = (DateTime)args[0];
            DateTime endDate = (DateTime)args[1];
            bool printHasRecordOnly = (bool)args[2];
            MemoryStream templateStream = (MemoryStream)args[3];
            string receiveName = (string)args[4];
            string receiveAddress = (string)args[5];
            string condName = (string)args[6];
            int condNumber = int.Parse((string)args[7]);
            bool IsInsertDate = (bool)args[8];
            bool printStudentList = (bool)args[9];
            bool printRemark = (bool)args[10];

            string reportName = "獎懲通知單(" + startDate.ToString("yyyy-MM-dd") + "至" + endDate.ToString("yyyy-MM-dd") + ")";

            ChengeDemerit(condName, condNumber);

            Dictionary<string, int> MDMapping = new Dictionary<string, int>();
            MDMapping.Add("大功", 0);
            MDMapping.Add("小功", 1);
            MDMapping.Add("嘉獎", 2);
            MDMapping.Add("大過", 3);
            MDMapping.Add("小過", 4);
            MDMapping.Add("警告", 5);

            int flag = 2;
            if (!string.IsNullOrEmpty(condName))
                flag = (MDMapping[condName] < 3) ? 1 : 0;

            MDFilter filter = new MDFilter();
            if (flag < 2)
                filter.SetCondition(MDMapping[condName], condNumber);
            #endregion

            #region 快取資訊

            //學生資訊
            Dictionary<string, Dictionary<string, string>> studentInfo = new Dictionary<string, Dictionary<string, string>>();

            //獎懲累計資料
            Dictionary<string, Dictionary<string, int>> studentDiscipline = new Dictionary<string, Dictionary<string, int>>();

            //獎懲明細
            Dictionary<string, List<string>> studentDisciplineDetail = new Dictionary<string, List<string>>();

            //所有學生ID
            List<string> allStudentID = new List<string>();

            //學生人數
            int currentStudentCount = 1;
            int totalStudentNumber = 0;

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

            #endregion

            #region 依據 ClassID 建立班級學生清單
            List<StudentRecord> classStudent = SelectedStudents;
            Dictionary<string, int> StudMeritSum = new Dictionary<string, int>();

            foreach (StudentRecord aStudent in classStudent)
            {
                string aStudentID = aStudent.ID;


                if (!StudMeritSum.ContainsKey(aStudentID))
                {
                    StudMeritSum.Add(aStudentID, 0);
                }

                if (!studentInfo.ContainsKey(aStudentID))
                    studentInfo.Add(aStudentID, new Dictionary<string, string>());

                TeacherRecord objT = aStudent.Class == null ? null : aStudent.Class.Teacher;

                studentInfo[aStudentID].Add("Name", aStudent.Name);
                studentInfo[aStudentID].Add("ClassName", aStudent.Class == null ? "" : aStudent.Class.Name);
                studentInfo[aStudentID].Add("SeatNo", aStudent.SeatNo);
                studentInfo[aStudentID].Add("StudentNumber", aStudent.StudentNumber);
                studentInfo[aStudentID].Add("Teacher", objT == null ? "" : objT.Name);

                if (!studentDiscipline.ContainsKey(aStudentID))
                    studentDiscipline.Add(aStudentID, new Dictionary<string, int>());
                if (!studentDisciplineDetail.ContainsKey(aStudentID))
                    studentDisciplineDetail.Add(aStudentID, new List<string>());

                if (!allStudentID.Contains(aStudentID))
                    allStudentID.Add(aStudentID);
            }
            #endregion

            #region 取得獎懲資料 日期區間
            DSXmlHelper helper = new DSXmlHelper("Request");
            helper.AddElement("Field");
            helper.AddElement("Field", "All");
            helper.AddElement("Condition");
            foreach (string var in allStudentID)
            {
                helper.AddElement("Condition", "RefStudentID", var);
            }

            if (IsInsertDate)
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
                string studentID = var.SelectSingleNode("RefStudentID").InnerText;

                DateTime occurDate = DateTime.Parse(var.SelectSingleNode("OccurDate").InnerText);
                string occurMonthDay = occurDate.Month + "/" + occurDate.Day;
                string reason = var.SelectSingleNode("Reason").InnerText;
                string remark = var.SelectSingleNode("Remark").InnerText;
                if (!studentDisciplineDetail.ContainsKey(studentID))
                    studentDisciplineDetail.Add(studentID, new List<string>());

                if (!studentDiscipline.ContainsKey(studentID))
                    studentDiscipline.Add(studentID, new Dictionary<string, int>());

                if (var.SelectSingleNode("MeritFlag").InnerText == "1")
                {
                    XmlElement meritElement = (XmlElement)var.SelectSingleNode("Detail/Discipline/Merit");
                    if (meritElement == null) continue;

                    if (MeritDemerit == "獎勵")
                    {
                        if (StudMeritSum.ContainsKey(studentID))
                        {
                            StudMeritSum[studentID] += GetMeritType(meritElement);
                        }
                    }

                    bool comma = false;
                    StringBuilder detailString = new StringBuilder("");
                    detailString.Append(occurMonthDay + " ");
                    if (!string.IsNullOrEmpty(reason))
                        detailString.Append(reason + " ");

                    foreach (string merit in meritTable.Keys)
                    {
                        int tryTimes;
                        int times = int.TryParse(meritElement.GetAttribute(meritTable[merit]), out tryTimes) ? tryTimes : 0;
                        if (times > 0)
                        {
                            if (!studentDiscipline[studentID].ContainsKey("Range" + merit))
                                studentDiscipline[studentID].Add("Range" + merit, 0);
                            studentDiscipline[studentID]["Range" + merit] += times;
                            if (comma)
                                detailString.Append(",");
                            detailString.Append(merit + times + "次");
                            comma = true;
                        }
                    }

                    if (printRemark)
                    {
                        if (!string.IsNullOrEmpty(remark))
                            detailString.Append(" (" + remark + ")");
                    }

                    studentDisciplineDetail[studentID].Add(detailString.ToString());

                }
                else if (var.SelectSingleNode("MeritFlag").InnerText == "0")
                {
                    XmlElement demeritElement = (XmlElement)var.SelectSingleNode("Detail/Discipline/Demerit");
                    if (demeritElement == null) continue;

                    bool cleared = false;
                    if (demeritElement.GetAttribute("Cleared") == "是")
                        cleared = true;

                    #region 懲戒比例換算 & 判斷

                    if (cleared == false && MeritDemerit == "懲戒")
                    {
                        if (StudMeritSum.ContainsKey(studentID))
                        {
                            StudMeritSum[studentID] += GetDemeritType(demeritElement);
                        }
                    }

                    #endregion

                    bool comma = false;
                    StringBuilder detailString = new StringBuilder("");
                    detailString.Append(occurMonthDay + " ");
                    if (!string.IsNullOrEmpty(reason))
                        detailString.Append(reason + " ");

                    foreach (string demerit in demeritTable.Keys)
                    {
                        int tryTimes;
                        int times = int.TryParse(demeritElement.GetAttribute(demeritTable[demerit]), out tryTimes) ? tryTimes : 0;
                        if (times > 0)
                        {
                            if (!studentDiscipline[studentID].ContainsKey("Range" + demerit))
                                studentDiscipline[studentID].Add("Range" + demerit, 0);
                            if (!cleared)
                            {
                                studentDiscipline[studentID]["Range" + demerit] += times;
                                if (comma)
                                    detailString.Append(",");
                                detailString.Append(demerit + times + "次");
                                comma = true;
                            }
                        }
                    }

                    if (printRemark)
                    {
                        if (!string.IsNullOrEmpty(remark))
                            detailString.Append(" (" + remark + ")");
                    }

                    if (!cleared)
                        studentDisciplineDetail[studentID].Add(detailString.ToString());
                }
            }
            #endregion

            #region 取得獎懲資料 學期累計
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
            helper.AddElement("Order");
            helper.AddElement("Order", "OccurDate", "asc");
            dsrsp = JHSchool.Compatibility.Feature.Student.QueryDiscipline.GetDiscipline(new DSRequest(helper));

            foreach (XmlElement var in dsrsp.GetContent().GetElements("Discipline"))
            {
                DateTime occurDate = DateTime.Parse(var.SelectSingleNode("OccurDate").InnerText);
                if (occurDate.CompareTo(endDate) == 1)
                    continue;

                string studentID = var.SelectSingleNode("RefStudentID").InnerText;

                if (!studentDiscipline.ContainsKey(studentID))
                    studentDiscipline.Add(studentID, new Dictionary<string, int>());

                if (var.SelectSingleNode("MeritFlag").InnerText == "1")
                {
                    XmlElement meritElement = (XmlElement)var.SelectSingleNode("Detail/Discipline/Merit");
                    if (meritElement == null) continue;

                    foreach (string merit in meritTable.Keys)
                    {
                        int tryTimes;
                        int times = int.TryParse(meritElement.GetAttribute(meritTable[merit]), out tryTimes) ? tryTimes : 0;
                        if (times > 0)
                        {
                            if (!studentDiscipline[studentID].ContainsKey("Semester" + merit))
                                studentDiscipline[studentID].Add("Semester" + merit, 0);
                            studentDiscipline[studentID]["Semester" + merit] += times;
                        }

                    }
                }
                else
                {
                    XmlElement demeritElement = (XmlElement)var.SelectSingleNode("Detail/Discipline/Demerit");
                    if (demeritElement == null) continue;

                    bool cleared = false;
                    if (demeritElement.GetAttribute("Cleared") == "是")
                        cleared = true;

                    foreach (string demerit in demeritTable.Keys)
                    {
                        int tryTimes;
                        int times = int.TryParse(demeritElement.GetAttribute(demeritTable[demerit]), out tryTimes) ? tryTimes : 0;
                        if (times > 0)
                        {
                            if (!studentDiscipline[studentID].ContainsKey("Semester" + demerit))
                                studentDiscipline[studentID].Add("Semester" + demerit, 0);
                            if (!cleared)
                                studentDiscipline[studentID]["Semester" + demerit] += times;
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

            #region 產生報表

            Aspose.Words.Document template = new Aspose.Words.Document(templateStream, "", Aspose.Words.LoadFormat.Doc, "");
            template.MailMerge.Execute(
                new string[] { "學校名稱", "學校地址", "學校電話" },
                new object[] { School.ChineseName, School.Address, School.Telephone }
                );

            Aspose.Words.Document doc = new Aspose.Words.Document();
            doc.RemoveAllChildren();

            Aspose.Words.Node sectionNode = template.Sections[0].Clone();

            totalStudentNumber = studentDiscipline.Count;

            foreach (string studentID in studentDiscipline.Keys)
            {
                if (printHasRecordOnly)
                {
                    if (studentDisciplineDetail[studentID].Count == 0)
                        continue;
                }

                if (MeritDemerit == "獎勵")
                {
                    if (StudMeritSum[studentID] < MaxMerit)
                        continue;
                }
                else if (MeritDemerit == "懲戒")
                {
                    if (StudMeritSum[studentID] < MaxDemerit)
                        continue;
                }
                else //未設定
                {

                }

                #region 過濾不需要列印的學生

                if (flag < 2)
                {
                    int A = 0, B = 0, C = 0;

                    if (flag == 1)
                    {
                        int tryM;
                        A = studentDiscipline[studentID].TryGetValue("Range大功", out tryM) ? tryM : 0;
                        B = studentDiscipline[studentID].TryGetValue("Range小功", out tryM) ? tryM : 0;
                        C = studentDiscipline[studentID].TryGetValue("Range嘉獎", out tryM) ? tryM : 0;
                    }
                    else if (flag == 0)
                    {
                        int tryD;
                        A = studentDiscipline[studentID].TryGetValue("Range大過", out tryD) ? tryD : 0;
                        B = studentDiscipline[studentID].TryGetValue("Range小過", out tryD) ? tryD : 0;
                        C = studentDiscipline[studentID].TryGetValue("Range警告", out tryD) ? tryD : 0;
                    }

                    //if (filter.IsFilter(A, B, C))
                    //    continue;
                }

                #endregion

                Aspose.Words.Document eachDoc = new Aspose.Words.Document();
                eachDoc.RemoveAllChildren();
                eachDoc.Sections.Add(eachDoc.ImportNode(sectionNode, true));

                //合併列印的資料
                Dictionary<string, object> mapping = new Dictionary<string, object>();
                Dictionary<string, string> eachStudentInfo = studentInfo[studentID];

                //學生資料
                mapping.Add("系統編號", "系統編號{" + studentID + "}");
                mapping.Add("學生姓名", BehTool.SurrogatePairString(eachStudentInfo["Name"]));
                mapping.Add("班級", eachStudentInfo["ClassName"]);
                mapping.Add("座號", eachStudentInfo["SeatNo"]);
                mapping.Add("學號", eachStudentInfo["StudentNumber"]);
                mapping.Add("導師", eachStudentInfo["Teacher"]);
                mapping.Add("資料期間", startDate.ToShortDateString() + " 至 " + endDate.ToShortDateString());

                //收件人資料
                if (receiveName == "監護人姓名")
                    mapping.Add("收件人姓名", BehTool.SurrogatePairString(eachStudentInfo["CustodianName"]));
                else if (receiveName == "父親姓名")
                    mapping.Add("收件人姓名", BehTool.SurrogatePairString(eachStudentInfo["FatherName"]));
                else if (receiveName == "母親姓名")
                    mapping.Add("收件人姓名", BehTool.SurrogatePairString(eachStudentInfo["MotherName"]));
                else
                    mapping.Add("收件人姓名", BehTool.SurrogatePairString(eachStudentInfo["Name"]));

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

                Dictionary<string, int> eachStudentDiscipline = studentDiscipline[studentID];

                //學生獎懲累計資料
                int count;
                mapping.Add("學期累計大功", eachStudentDiscipline.TryGetValue("Semester大功", out count) ? "" + count : "0");
                mapping.Add("學期累計小功", eachStudentDiscipline.TryGetValue("Semester小功", out count) ? "" + count : "0");
                mapping.Add("學期累計嘉獎", eachStudentDiscipline.TryGetValue("Semester嘉獎", out count) ? "" + count : "0");
                mapping.Add("學期累計大過", eachStudentDiscipline.TryGetValue("Semester大過", out count) ? "" + count : "0");
                mapping.Add("學期累計小過", eachStudentDiscipline.TryGetValue("Semester小過", out count) ? "" + count : "0");
                mapping.Add("學期累計警告", eachStudentDiscipline.TryGetValue("Semester警告", out count) ? "" + count : "0");
                mapping.Add("本期累計大功", eachStudentDiscipline.TryGetValue("Range大功", out count) ? "" + count : "0");
                mapping.Add("本期累計小功", eachStudentDiscipline.TryGetValue("Range小功", out count) ? "" + count : "0");
                mapping.Add("本期累計嘉獎", eachStudentDiscipline.TryGetValue("Range嘉獎", out count) ? "" + count : "0");
                mapping.Add("本期累計大過", eachStudentDiscipline.TryGetValue("Range大過", out count) ? "" + count : "0");
                mapping.Add("本期累計小過", eachStudentDiscipline.TryGetValue("Range小過", out count) ? "" + count : "0");
                mapping.Add("本期累計警告", eachStudentDiscipline.TryGetValue("Range警告", out count) ? "" + count : "0");

                //獎懲明細
                object[] objectValues = new object[] { studentDisciplineDetail[studentID] };
                mapping.Add("獎懲明細", objectValues);

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
            if (printStudentList)
            {
                if (printHasRecordOnly)
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
                        if (studentDisciplineDetail[each].Count == 0)
                            continue;

                        if (MeritDemerit == "獎勵")
                        {
                            if (StudMeritSum[each] < MaxMerit)
                                continue;
                        }
                        else if (MeritDemerit == "懲戒")
                        {
                            if (StudMeritSum[each] < MaxDemerit)
                                continue;
                        }
                        else //未設定
                        {

                        }

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
            }
            #endregion

            string path = Path.Combine(Application.StartupPath, "Reports");
            string path2 = Path.Combine(Application.StartupPath, "Reports");

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");
            path2 = Path.Combine(path2, reportName + "(學生清單).xlsx");

            string message = "【電子報表通知】您好 本期「{0}」已產生,可於電子報表中檢視「資料期間：{1} 至 {2}」";
            e.Result = new object[] { reportName, path, doc, path2, printStudentList, wb , string.Format(message, "獎懲通知單", startDate.ToShortDateString(), endDate.ToShortDateString()) };
        }

        private void DisciplineNotification_MailMerge_MergeField(object sender, Aspose.Words.Reporting.MergeFieldEventArgs e)
        {
            #region MailMerge_MergeField
            if (e.FieldName == "獎懲明細")
            {
                object[] objectValues = (object[])e.FieldValue;
                List<string> eachStudentDisciplineDetail = (List<string>)objectValues[0];

                Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(e.Document);

                builder.MoveToField(e.Field, false);
                builder.StartTable();
                builder.CellFormat.ClearFormatting();
                builder.CellFormat.Borders.ClearFormatting();
                builder.CellFormat.VerticalAlignment = Aspose.Words.CellVerticalAlignment.Center;
                builder.CellFormat.LeftPadding = 3.0;
                builder.RowFormat.LeftIndent = 0.0;
                builder.RowFormat.Height = 15.0;

                int rowNumber = 4;
                if (eachStudentDisciplineDetail.Count > rowNumber * 2)
                {
                    rowNumber = eachStudentDisciplineDetail.Count / 2;
                    rowNumber += eachStudentDisciplineDetail.Count % 2;
                }

                if (eachStudentDisciplineDetail.Count > rowNumber * 2)
                {
                    rowNumber += (eachStudentDisciplineDetail.Count - (rowNumber * 2)) / 2;
                    rowNumber += (eachStudentDisciplineDetail.Count - (rowNumber * 2)) % 2;
                }

                for (int j = 0; j < rowNumber; j++)
                {
                    builder.InsertCell();
                    builder.CellFormat.Borders.Right.LineStyle = Aspose.Words.LineStyle.Single;
                    builder.CellFormat.Borders.Right.Color = Color.Black;
                    if (j < eachStudentDisciplineDetail.Count)
                        builder.Write(eachStudentDisciplineDetail[j]);
                    builder.InsertCell();
                    if (j + rowNumber < eachStudentDisciplineDetail.Count)
                        builder.Write(eachStudentDisciplineDetail[j + rowNumber]);
                    builder.EndRow();
                }

                builder.EndTable();

                e.Text = string.Empty;
            }
            #endregion
        }

        class MDFilter
        {
            #region 功過條件過濾
            private int _condName = 0; //0:A, 1:B, 2:C
            private int _condNumber = 0;

            public MDFilter()
            {
            }

            public void SetCondition(int name, int number)
            {
                _condName = name;
                _condNumber = number;
            }

            public bool IsFilter(int A, int B, int C)
            {
                bool filtered = false;

                switch (_condName)
                {
                    case 5:
                    case 2:
                        if ((A + B) > 0)
                            filtered = false;
                        else if (C >= _condNumber)
                            filtered = false;
                        else
                            filtered = true;
                        break;
                    case 4:
                    case 1:
                        if (A > 0)
                            filtered = false;
                        else if (B >= _condNumber)
                            filtered = false;
                        else
                            filtered = true;
                        break;
                    case 3:
                    case 0:
                        if (A >= _condNumber)
                            filtered = false;
                        else
                            filtered = true;
                        break;
                    default:
                        break;
                }

                return filtered;
            }
            #endregion
        }
    }
}
