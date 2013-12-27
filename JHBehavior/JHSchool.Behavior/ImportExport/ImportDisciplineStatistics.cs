using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using JHSchool.Data;
using SmartSchool.API.PlugIn;

namespace JHSchool.Behavior.ImportExport
{
    class ImportDisciplineStatistics : SmartSchool.API.PlugIn.Import.Importer
    {
        private List<string> Keys = new List<string>();

        public ImportDisciplineStatistics()
        {
            this.Image = null;
            this.Text = "匯入獎勵懲戒統計";
        }

        public override void InitializeImport(SmartSchool.API.PlugIn.Import.ImportWizard wizard)
        {
            Dictionary<string, JHMoralScoreRecord> CacheMoralScore = new Dictionary<string, JHMoralScoreRecord>();
            List<JHDisciplineRecord> DiscipleRecords = new List<JHDisciplineRecord>();

            wizard.RequiredFields.AddRange("學年度","學期");
            wizard.ImportableFields.AddRange("學年度", "學期", "大功", "小功", "嘉獎", "大過", "小過", "警告");
            wizard.PackageLimit = 1000;
            wizard.ValidateStart += (sender,e) =>
            {
                Keys.Clear();
            };

            wizard.ValidateRow += (sender,e)=>
            {
                int schoolYear, semester;
                #region 驗共同必填欄位
                if (!int.TryParse(e.Data["學年度"], out schoolYear))
                {
                    e.ErrorFields.Add("學年度", "必需輸入數字");
                }
                if (!int.TryParse(e.Data["學期"], out semester))
                {
                    e.ErrorFields.Add("學期", "必需輸入數字");
                }
                #endregion
                int awardA = 0;
                int awardB = 0;
                int awardC = 0;
                int faultA = 0;
                int faultB = 0;
                int faultC = 0;
                #region 驗證必填欄位
                if (!int.TryParse(e.Data["大功"], out awardA))
                {
                    e.ErrorFields.Add("大功", "必需輸入數字");
                }
                if (!int.TryParse(e.Data["小功"], out awardB))
                {
                    e.ErrorFields.Add("小功", "必需輸入數字");
                }
                if (!int.TryParse(e.Data["嘉獎"], out awardC))
                {
                    e.ErrorFields.Add("嘉獎", "必需輸入數字");
                }
                if (!int.TryParse(e.Data["大過"], out faultA))
                {
                    e.ErrorFields.Add("大過", "必需輸入數字");
                }
                if (!int.TryParse(e.Data["小過"], out faultB))
                {
                    e.ErrorFields.Add("小過", "必需輸入數字");
                }
                if (!int.TryParse(e.Data["警告"], out faultC))
                {
                    e.ErrorFields.Add("警告", "必需輸入數字");
                }
                #endregion
                #region 驗證主鍵

                string Key = e.Data.ID + "-" + e.Data["學年度"] + "-" + e.Data["學期"];

                if (Keys.Contains(Key))
                    e.ErrorMessage = "學生編號、學年度、學期之組合不能重覆!";
                else
                    Keys.Add(Key);
                #endregion
            };
            //匯入資料完成的事件
            wizard.ImportComplete += (sender, e) => MessageBox.Show("匯入完成!");
            //實際匯入資料完成的事件
            wizard.ImportPackage += (sender,e)=>
            {
                //取得學生的所有日常生活表現紀錄列表
                foreach (JHMoralScoreRecord record in JHMoralScore.SelectByStudentIDs(e.Items.Select(x => x.ID)))
                    if (!CacheMoralScore.ContainsKey(record.ID))
                        CacheMoralScore.Add(record.ID, record);

                //取得學生的所有獎勵紀錄列表
                List<JHMeritRecord> meritrecords = JHMerit.SelectByStudentIDs(e.Items.Select(x => x.ID));
                
                //取得學生的所有懲戒紀錄列表
                List<JHDemeritRecord> demeritrecords = JHDemerit.SelectByStudentIDs(e.Items.Select(x => x.ID));

                //要更新的德行成績列表
                List<JHMoralScoreRecord> updateMoralScores = new List<JHMoralScoreRecord>();

                //要新增的德行成績列表
                List<JHMoralScoreRecord> insertMoralScores = new List<JHMoralScoreRecord>();

                //取得每筆匯入資料
                foreach (RowData row in e.Items)
                {
                    int schoolYear = K12.Data.Int.Parse(row["學年度"]);
                    int semester = K12.Data.Int.Parse(row["學期"]);

                    //根據學生編號、學年度及學期尋找是否有對應的德行成績
                    List<JHMoralScoreRecord> records = CacheMoralScore.Values.Where(x => x.RefStudentID.Equals(row.ID) && (x.SchoolYear == schoolYear) && (x.Semester == semester)).ToList();

                    #region 計算學生獎勵及懲戒明細統計
                    //根據學年度學期取得獎勵紀錄
                    List<JHMeritRecord> merits = meritrecords.Where(x => x.RefStudentID == row.ID && x.SchoolYear == schoolYear && x.Semester == semester).ToList();
                    
                    //根據學年度學期取得懲戒紀錄
                    List<JHDemeritRecord> demerits = demeritrecords.Where(x => x.RefStudentID == row.ID && x.SchoolYear == schoolYear && x.Semester == semester).ToList();

                    int MeritA = 0;
                    int MeritB = 0;
                    int MeritC = 0;
                    int DemeritA = 0;
                    int DemeritB = 0;
                    int DemeritC = 0;

                    foreach (JHMeritRecord merit in merits)
                    {
                        MeritA += merit.MeritA.HasValue ? merit.MeritA.Value : 0;
                        MeritB += merit.MeritB.HasValue ? merit.MeritB.Value : 0;
                        MeritC += merit.MeritC.HasValue ? merit.MeritC.Value : 0;
                    }

                    foreach (JHDemeritRecord demerit in demerits)
                    {
                        //必需要沒有銷過的才列入統計
                        if (demerit.Cleared.Equals(string.Empty))
                        {
                            DemeritA += demerit.DemeritA.HasValue ? demerit.DemeritA.Value : 0;
                            DemeritB += demerit.DemeritB.HasValue ? demerit.DemeritB.Value : 0;
                            DemeritC += demerit.DemeritC.HasValue ? demerit.DemeritC.Value : 0;
                        }
                    }
                    #endregion

                    //該學生的學年度及學期德行成績已存在
                    if (records.Count > 0)
                    {
                        //根據學生編號、學年度、學期及日期取得的缺曠記錄應該只有一筆
                        JHMoralScoreRecord record = records[0];

                        //確認Sumarry的結構存在
                        MakeSureElement(record);

                        //更新對應的獎懲值
                        if (record.InitialSummary != null)
                        {
                            XmlNode NodeMeritA = record.InitialSummary.SelectSingleNode("DisciplineStatistics/Merit/@A");

                            NodeMeritA.InnerText = ""+(K12.Data.Int.Parse(GetUpdateFieldValue(row, "大功", NodeMeritA.InnerText))-MeritA);

                            XmlNode NodeMeritB = record.InitialSummary.SelectSingleNode("DisciplineStatistics/Merit/@B");

                            NodeMeritB.InnerText = ""+(K12.Data .Int.Parse(GetUpdateFieldValue(row, "小功", NodeMeritB.InnerText))-MeritB);

                            XmlNode NodeMeritC = record.InitialSummary.SelectSingleNode("DisciplineStatistics/Merit/@C");

                            NodeMeritC.InnerText = ""+(K12.Data.Int.Parse(GetUpdateFieldValue(row, "嘉獎", NodeMeritC.InnerText))-MeritC);

                            XmlNode NodeDemeritA = record.InitialSummary.SelectSingleNode("DisciplineStatistics/Demerit/@A");

                            NodeDemeritA.InnerText = "" + (K12.Data.Int.Parse(GetUpdateFieldValue(row, "大過", NodeDemeritA.InnerText)) - DemeritA);

                            XmlNode NodeDemeritB = record.InitialSummary.SelectSingleNode("DisciplineStatistics/Demerit/@B");

                            NodeDemeritB.InnerText = ""+(K12.Data.Int.Parse(GetUpdateFieldValue(row, "小過", NodeDemeritB.InnerText))-DemeritB);

                            XmlNode NodeDemeritC = record.InitialSummary.SelectSingleNode("DisciplineStatistics/Demerit/@C");

                            NodeDemeritC.InnerText = ""+(K12.Data.Int.Parse(GetUpdateFieldValue(row, "警告", NodeDemeritC.InnerText))-DemeritC);

                            updateMoralScores.Add(record);
                        }
                    }
                    else
                    {
                        JHMoralScoreRecord record = new JHMoralScoreRecord();

                        record.SchoolYear = schoolYear;
                        record.Semester = semester;
                        record.RefStudentID = row.ID;

                        MakeSureElement(record);

                        string NodeMeritA = ""+(K12.Data.Int.Parse(GetInsertFieldValue(row, "大功"))-MeritA);
                        string NodeMeritB = "" + (K12.Data.Int.Parse(GetInsertFieldValue(row, "小功")) - MeritB);
                        string NodeMeritC = "" + (K12.Data.Int.Parse(GetInsertFieldValue(row, "嘉獎")) - MeritC);
                        string NodeDemeritA = "" + (K12.Data.Int.Parse(GetInsertFieldValue(row, "大過")) - DemeritA);
                        string NodeDemeritB = "" + (K12.Data.Int.Parse(GetInsertFieldValue(row, "小過")) - DemeritB);
                        string NodeDemeritC = "" + (K12.Data.Int.Parse(GetInsertFieldValue(row, "警告")) - DemeritC);

                        record.InitialSummary.SelectSingleNode("DisciplineStatistics/Merit/@A").InnerText = NodeMeritA;
                        record.InitialSummary.SelectSingleNode("DisciplineStatistics/Merit/@B").InnerText = NodeMeritB;
                        record.InitialSummary.SelectSingleNode("DisciplineStatistics/Merit/@C").InnerText = NodeMeritC;

                        record.InitialSummary.SelectSingleNode("DisciplineStatistics/Demerit/@A").InnerText = NodeDemeritA;
                        record.InitialSummary.SelectSingleNode("DisciplineStatistics/Demerit/@B").InnerText = NodeDemeritB;
                        record.InitialSummary.SelectSingleNode("DisciplineStatistics/Demerit/@C").InnerText = NodeDemeritC;

                        insertMoralScores.Add(record);
                    }
                }

                if (updateMoralScores.Count > 0)
                {
                    //foreach (JHMoralScoreRecord record in updateMoralScores)
                    //    RemoveEmptyDisciplineElement(record);
                    JHMoralScore.Update(updateMoralScores);
                }
                if (insertMoralScores.Count > 0)
                {
                    //foreach (JHMoralScoreRecord record in insertMoralScores)
                    //    RemoveEmptyDisciplineElement(record);
                    JHMoralScore.Insert(insertMoralScores);
                }
            };
        }

        private void RemoveEmptyDisciplineElement(JHMoralScoreRecord record)
        {
            XmlElement MeritElement = record.InitialSummary.SelectSingleNode("DisciplineStatistics/Merit") as XmlElement;

            record.InitialSummary.SelectSingleNode("DisciplineStatistics").RemoveChild(MeritElement);

            XmlElement DemeritElement = record.InitialSummary.SelectSingleNode("DisciplineStatistics/Demerit") as XmlElement;

            record.InitialSummary.SelectSingleNode("DisciplineStatistics").RemoveChild(DemeritElement);
        }

        //針對新增獎懲統計取得值的函式，儘量幫其補0
        public string GetInsertFieldValue(RowData Row,string FieldName)
        {
            
            if (Row.ContainsKey(FieldName))
                if (string.IsNullOrEmpty(FieldName))
                    return "0"; //假設使用者的值為空白，幫其補0
                else
                    return Row[FieldName];
            else //假設使用者沒有選此欄位，幫其補0
                return "0";
        }

        //針對更新獎懲統計取得值的函式，儘量幫其補0
        public string GetUpdateFieldValue(RowData Row, string FieldName,string CurrentValue)
        {
            if (Row.ContainsKey(FieldName))
                if (string.IsNullOrEmpty(Row[FieldName]))
                    return "0";
                else
                    return Row[FieldName];
            else
                return string.IsNullOrEmpty(CurrentValue) ? "0" : CurrentValue;
        }

        public void MakeSureElement(JHMoralScoreRecord record)
        {
            if (record.InitialSummary == null)
            {
                XmlDocument xmldoc = new XmlDocument();
                xmldoc.LoadXml("<InitialSummary/>");
                record.InitialSummary = xmldoc.DocumentElement;
            }

            if (record.InitialSummary.SelectSingleNode("DisciplineStatistics") == null)
            {
               XmlElement Element = record.InitialSummary.OwnerDocument.CreateElement("DisciplineStatistics");
               record.InitialSummary.AppendChild(Element);
            }

            if (record.InitialSummary.SelectSingleNode("DisciplineStatistics/Merit") == null)
            {
               XmlDocumentFragment Fragment = record.InitialSummary.OwnerDocument.CreateDocumentFragment();
               Fragment.InnerXml = "<Merit A=\"\" B=\"\" C=\"\"/>";
               record.InitialSummary.SelectSingleNode("DisciplineStatistics").AppendChild(Fragment);
            }

            if (record.InitialSummary.SelectSingleNode("DisciplineStatistics/Demerit") == null)
            {
               XmlDocumentFragment Fragment = record.InitialSummary.OwnerDocument.CreateDocumentFragment();
               Fragment.InnerXml = "<Demerit A=\"\" B=\"\" C=\"\"/>";
               record.InitialSummary.SelectSingleNode("DisciplineStatistics").AppendChild(Fragment);
            }           
        }
    }
}