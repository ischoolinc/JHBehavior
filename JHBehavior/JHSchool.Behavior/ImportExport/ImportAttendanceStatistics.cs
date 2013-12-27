using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using JHSchool.Behavior.BusinessLogic;
using JHSchool.Data;
using SmartSchool.API.PlugIn;

namespace JHSchool.Behavior.ImportExport
{
    class ImportAttendanceStatistics : SmartSchool.API.PlugIn.Import.Importer
    {
        List<string> absenceList = new List<string>();
        List<string> periodtypelist = new List<string>();
        List<string> Keys = new List<string>();

        public ImportAttendanceStatistics()
        {
            this.Image = null;
            this.Text = "匯入缺曠統計";
        }

        public override void InitializeImport(SmartSchool.API.PlugIn.Import.ImportWizard wizard)
        {
            wizard.PackageLimit = 1000;
            //缺曠記錄可匯入的欄位
            wizard.ImportableFields.AddRange("學年度", "學期", "節次類型", "缺曠假別", "缺曠統計值");
            //缺曠記錄必需要有的欄位
            wizard.RequiredFields.AddRange("學年度", "學期", "節次類型", "缺曠假別", "缺曠統計值");
            //開始驗證時的事件
            wizard.ValidateStart += new System.EventHandler<SmartSchool.API.PlugIn.Import.ValidateStartEventArgs>(wizard_ValidateStart);
            //驗證每行資料的事件
            wizard.ValidateRow += new System.EventHandler<SmartSchool.API.PlugIn.Import.ValidateRowEventArgs>(wizard_ValidateRow);
            //實際匯入資料的事件
            wizard.ImportPackage += new System.EventHandler<SmartSchool.API.PlugIn.Import.ImportPackageEventArgs>(wizard_ImportPackage);
            //匯入資料完成的事件
            wizard.ImportComplete += (sender, e) => MessageBox.Show("匯入完成!");
        }

        void wizard_ValidateStart(object sender, SmartSchool.API.PlugIn.Import.ValidateStartEventArgs e)
        {
            //取得系統假別對照表
            absenceList = JHAbsenceMapping.SelectAll().Select(x => x.Name).ToList();

            //取得系統節次型態列表
            periodtypelist = JHPeriodMapping.SelectAll().Select(x => x.Type).ToList();

            Keys.Clear();
        }

        void wizard_ValidateRow(object sender, SmartSchool.API.PlugIn.Import.ValidateRowEventArgs e)
        {
            //"學年度", "學期", "日期", "缺曠假別", "缺曠節次"
            #region 驗各欄位填寫格式
            int t;
            foreach (string field in e.SelectFields)
            {
                string value = e.Data[field];
                switch (field)
                {
                    default:
                        break;
                    case "學年度":
                        if (value == "" || !int.TryParse(value, out t))
                            e.ErrorFields.Add(field, "此欄為必填欄位，必須填入整數。");
                        break;
                    case "學期":
                        if (value == "" || !int.TryParse(value, out t))
                        {
                            e.ErrorFields.Add(field, "此欄為必填欄位，必須填入整數。");
                        }
                        else if (t != 1 && t != 2)
                        {
                            e.ErrorFields.Add(field, "必須填入1或2");
                        }
                        break;
                    case "節次類型":
                        if (!periodtypelist.Contains(value))
                        {
                            e.ErrorFields.Add(field, "輸入值" + value + "不在節次類型清單中。");
                        }
                        break;
                    case "缺曠假別":
                        if (!absenceList.Contains(value))
                        {
                            e.ErrorFields.Add(field, "輸入值" + value + "不在假別清單中。");
                        }
                        break;
                    case "缺曠統計值":
                        if (value == "" || !int.TryParse(value, out t))
                        {
                            e.ErrorFields.Add(field, "此欄為必填欄位");
                        }
                        break;
                }
            }
            #endregion
            #region 驗證主鍵

            string Key = e.Data.ID + "-" + e.Data["學年度"] + "-" + e.Data["學期"] + "-" + e.Data["節次類型"] + "-" + e.Data["缺曠假別"];
            string errorMessage = string.Empty;

            if (Keys.Contains(Key))
                errorMessage = "學生編號、學年度、學期、節次類型及缺曠假別之組合不能重覆!";
            else
                Keys.Add(Key);

            e.ErrorMessage = errorMessage;

            #endregion
        }

        void wizard_ImportPackage(object sender, SmartSchool.API.PlugIn.Import.ImportPackageEventArgs e)
        {
            //以學生編號、學年度、學期做為鍵值
            List<string> keyList = new List<string>();
            //鍵值所對應到的學年度
            Dictionary<string, int> schoolYearMapping = new Dictionary<string, int>();
            //鍵值所對應到的學期
            Dictionary<string, int> semesterMapping = new Dictionary<string, int>();
            //鍵值所對應到的學生編號
            Dictionary<string, string> studentIDMapping = new Dictionary<string, string>();
            //鍵值所對應到的資料列表
            Dictionary<string, List<RowData>> rowsMapping = new Dictionary<string, List<RowData>>();
            //鍵值所對應到的學生日常生活表現列表
            Dictionary<string, List<JHMoralScoreRecord>> studentMoralScoreInfo = new Dictionary<string, List<JHMoralScoreRecord>>();
            //取得學生的所有缺曠紀錄列表
            List<JHAttendanceRecord> attendancerecords = JHAttendance.SelectByStudentIDs(e.Items.Select(x => x.ID));

            //掃描每行資料，定出資料的PrimaryKey，並且將PrimaryKey對應到的資料寫成Dictionary
            foreach (RowData Row in e.Items)
            {
                #region 取得學年度、學期及學生編號並組成鍵值
                int schoolYear = int.Parse(Row["學年度"]);
                int semester = int.Parse(Row["學期"]);
                string studentID = Row.ID;
                string key = schoolYear + "^_^" + semester + "^_^" + studentID;
                #endregion

                if (!keyList.Contains(key))
                {
                    keyList.Add(key);
                    schoolYearMapping.Add(key, schoolYear);
                    semesterMapping.Add(key, semester);
                    studentIDMapping.Add(key, studentID);
                    rowsMapping.Add(key, new List<RowData>());
                }
                rowsMapping[key].Add(Row);
            }

            #region 取得學生日常生活表現紀錄
            foreach (JHMoralScoreRecord var in JHMoralScore.SelectByStudentIDs(studentIDMapping.Values.Distinct()))
            {
                if (!studentMoralScoreInfo.ContainsKey(var.RefStudentID))
                    studentMoralScoreInfo.Add(var.RefStudentID, new List<JHMoralScoreRecord>());
                studentMoralScoreInfo[var.RefStudentID].Add(var);
            }
            #endregion

            List<JHMoralScoreRecord> InsertRecords = new List<JHMoralScoreRecord>();
            List<JHMoralScoreRecord> UpdateRecords = new List<JHMoralScoreRecord>();

            foreach (string key in keyList)
            {
                #region 根據學生編號、學年度、學期及日期取得日常生活表現紀錄
                List<JHMoralScoreRecord> records = new List<JHMoralScoreRecord>();
               
                if (studentMoralScoreInfo.ContainsKey(studentIDMapping[key]))
                    records = studentMoralScoreInfo[studentIDMapping[key]].Where(x => x.SchoolYear == schoolYearMapping[key] && x.Semester == semesterMapping[key]).ToList();
                #endregion

                //根據鍵值取得匯入資料，該匯入資料應該是有相同的學生編號、學年度、學期及缺曠日期
                List<RowData> Rows = rowsMapping[key];

                //根據學生編號、學年度、學期取得缺曠紀錄，並且統計值
                List<AbsenceCountRecord> absencecounts = AutoSummary.Calculate(attendancerecords.Where(x => x.RefStudentID == studentIDMapping[key] && x.SchoolYear == schoolYearMapping[key] && x.Semester == semesterMapping[key]).ToList());

                //該筆缺曠記錄已存在系統中
                if (records.Count > 0)
                {
                    //根據學生編號、學年度、學期及日期取得的缺曠記錄應該只有一筆
                    JHMoralScoreRecord record = records[0];

                    for (int i = 0; i < Rows.Count; i++)
                    {
                        //取得匯入資料的假別
                        string Absence = Rows[i]["缺曠假別"];
                        string PeriodType = Rows[i]["節次類型"];
                        string AbsenceCount = Rows[i]["缺曠統計值"];

                        List<AbsenceCountRecord> AbsenceCounts = absencecounts.Where(x => x.PeriodType == PeriodType && x.Name == Absence).ToList();

                        //將匯入值減去明細
                        if (AbsenceCounts.Count > 0)
                            AbsenceCount = "" + (K12.Data.Int.Parse(AbsenceCount) - AbsenceCounts[0].Count);

                        bool IsExist = false;

                        MakeSureElement(record);

                        //節次已經存在會更新假別
                        //Absence的結構：<Absence Count="3" Name="曠課" PeriodType="集會" />
                        foreach (XmlNode CurrentAbsence in record.InitialSummary.SelectNodes("AttendanceStatistics/Absence"))
                        {
                            XmlElement Element = CurrentAbsence as XmlElement;

                            //根據相同的缺曠類別及節次類別來更新缺曠次數
                            if (Element!=null)
                                if (Element.GetAttribute("Name").Equals(Absence) && Element.GetAttribute("PeriodType").Equals(PeriodType))
                                {
                                    Element.SetAttribute("Count",AbsenceCount);
                                    IsExist = true;
                                }
                        }

                        //若是節次不存在會根據節次及假別新增缺曠明細
                        if (!IsExist)
                        {
                            XmlElement Element = record.InitialSummary.OwnerDocument.CreateElement("Absence");

                            Element.SetAttribute("Count",AbsenceCount);
                            Element.SetAttribute("Name",Absence);
                            Element.SetAttribute("PeriodType",PeriodType);

                            record.InitialSummary.SelectSingleNode("AttendanceStatistics").AppendChild(Element);
                        }
                    }

                    UpdateRecords.Add(record);
                }
                else //該筆缺曠記錄沒有存在系統中
                {
                    JHMoralScoreRecord record = new JHMoralScoreRecord();

                    record.SchoolYear = schoolYearMapping[key];
                    record.Semester = semesterMapping[key];
                    record.RefStudentID = Rows[0].ID;

                    //確認在記錄中有正確的XML結構
                    //<Summary><AttendanceStatistics/><DisciplineStatistics/></Summary>
                    MakeSureElement(record);

                    //將屬於同樣一筆的匯入資料都加入到同樣的缺曠記錄中的明細
                    foreach (RowData Row in Rows)
                    {
                        if (!string.IsNullOrEmpty(Row["缺曠統計值"]))
                        {
                            //取得匯入資料的假別
                            string Absence = Row["缺曠假別"];
                            string PeriodType = Row["節次類型"];
                            string AbsenceCount = Row["缺曠統計值"];

                            List<AbsenceCountRecord> AbsenceCounts = absencecounts.Where(x => x.PeriodType == PeriodType && x.Name == Absence).ToList();

                            //將統計值減去明細
                            if (AbsenceCounts.Count > 0)
                                AbsenceCount = "" + (K12.Data.Int.Parse(AbsenceCount) - AbsenceCounts[0].Count);

                            XmlElement Element = record.InitialSummary.OwnerDocument.CreateElement("Absence");

                            Element.SetAttribute("Count", AbsenceCount);
                            Element.SetAttribute("Name", Absence);
                            Element.SetAttribute("PeriodType", PeriodType);
                            record.InitialSummary.SelectSingleNode("AttendanceStatistics").AppendChild(Element);
                        }
                    }

                    InsertRecords.Add(record);
                }
            }

            if (InsertRecords.Count > 0)
            {
                //為了刪除狀況特殊加的程式
                //foreach (JHMoralScoreRecord record in InsertMoralScores)
                //    RemoveElement(record); 

                JHMoralScore.Insert(InsertRecords);
            }

            if (UpdateRecords.Count > 0)
            {
                //更新時移除空白的統計值
                foreach (JHMoralScoreRecord record in UpdateRecords)
                {
                    List<XmlNode> RemoveNodes = new List<XmlNode>();

                    foreach (XmlNode CurrentAbsence in record.InitialSummary.SelectNodes("AttendanceStatistics/Absence"))
                    {
                        XmlElement Element = CurrentAbsence as XmlElement;

                        if (Element != null)
                            if (string.IsNullOrEmpty(Element.GetAttribute("Count")))
                                RemoveNodes.Add(CurrentAbsence);
                    }

                    foreach (XmlNode Node in RemoveNodes)
                        record.InitialSummary.SelectSingleNode("AttendanceStatistics").RemoveChild(Node);
                }

                //為了刪除狀況特殊加的程式
                //foreach (JHMoralScoreRecord record in UpdateMoralScores)
                //    RemoveElement(record); 
                
                JHMoralScore.Update(UpdateRecords);
            }
        }

        /// <summary>
        /// 移除缺曠統計節點
        /// </summary>
        /// <param name="record"></param>
        private void RemoveElement(JHMoralScoreRecord record)
        {
            XmlElement Element = record.InitialSummary.SelectSingleNode("AttendanceStatistics") as XmlElement;

            Element.RemoveAll();
        }

        /// <summary>
        /// 確認在記錄中有正確的XML結構 
        /// </summary>
        /// <param name="record"></param>
        public void MakeSureElement(JHMoralScoreRecord record)
        {
            //<Summary><AttendanceStatistics/><DisciplineStatistics/></Summary>
            if (record.InitialSummary == null)
            {
                XmlDocument xmldoc = new XmlDocument();
                xmldoc.LoadXml("<InitialSummary/>");
                record.InitialSummary = xmldoc.DocumentElement;
            }

            if (record.InitialSummary.SelectSingleNode("AttendanceStatistics") == null)
            {
                XmlElement Element = record.InitialSummary.OwnerDocument.CreateElement("AttendanceStatistics");
                record.InitialSummary.AppendChild(Element);
            }
        }
    }
}