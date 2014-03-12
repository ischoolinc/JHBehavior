using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using JHSchool.Data;
using SmartSchool.API.PlugIn;
using System.Text;

namespace JHSchool.Behavior.ImportExport
{
    class ImportDiscipline : SmartSchool.API.PlugIn.Import.Importer
    {
        public ImportDiscipline()
        {
            this.Image = null;
            this.Text = "匯入獎勵懲戒記錄";
        }

        public override void InitializeImport(SmartSchool.API.PlugIn.Import.ImportWizard wizard)
        {
            Dictionary<string, JHDisciplineRecord> CacheDiscipline = new Dictionary<string, JHDisciplineRecord>();

            VirtualRadioButton chose1 = new VirtualRadioButton("比對事由變更獎懲次數", false);
            VirtualRadioButton chose2 = new VirtualRadioButton("比對獎懲次數變更事由", false);
            chose1.CheckedChanged += delegate
            {
                if (chose1.Checked)
                {
                    wizard.RequiredFields.Clear();
                    wizard.RequiredFields.AddRange("學年度", "學期", "日期", "事由");
                }
            };
            chose2.CheckedChanged += delegate
            {
                if (chose2.Checked)
                {
                    wizard.RequiredFields.Clear();
                    wizard.RequiredFields.AddRange("學年度", "學期", "日期", "大功", "小功", "嘉獎", "大過", "小過", "警告");
                }
            };
            wizard.ImportableFields.AddRange("學年度", "學期", "日期", "地點", "大功", "小功", "嘉獎", "大過", "小過", "警告", "事由", "是否銷過", "銷過日期", "銷過事由", "登錄日期");
            wizard.Options.AddRange(chose1, chose2);
            chose1.Checked = true;
            wizard.PackageLimit = 1000;
            bool allPass = true;
            int insertRecords = 0;
            int updataRecords = 0;

            wizard.ValidateStart += delegate(object sender, SmartSchool.API.PlugIn.Import.ValidateStartEventArgs e)
            {
                foreach (JHDisciplineRecord record in JHDiscipline.SelectByStudentIDs(e.List))
                    if (!CacheDiscipline.ContainsKey(record.ID))
                        CacheDiscipline.Add(record.ID, record);

                allPass = true;
            };

            wizard.ValidateRow += delegate(object sender, SmartSchool.API.PlugIn.Import.ValidateRowEventArgs e)
            {
                #region 驗證資料
                bool pass = true;
                int schoolYear, semester;
                DateTime occurdate;
                bool isInsert = false;
                bool isUpdata = false;
                #region 驗共同必填欄位
                if (!int.TryParse(e.Data["學年度"], out schoolYear))
                {
                    e.ErrorFields.Add("學年度", "必需輸入數字");
                    pass = false;
                }
                if (!int.TryParse(e.Data["學期"], out semester))
                {
                    e.ErrorFields.Add("學期", "必需輸入數字");
                    pass = false;
                }
                if (!DateTime.TryParse(e.Data["日期"], out occurdate))
                {
                    e.ErrorFields.Add("日期", "輸入格式為 西元年//月//日");
                    pass = false;
                }
                #endregion
                if (!pass)
                {
                    allPass = false;
                    return;
                }
                if (chose1.Checked)
                {
                    #region 以事由為Key更新
                    string reason = e.Data["事由"];
                    int match = 0;
                    foreach (JHDisciplineRecord rewardInfo in CacheDiscipline.Values.Where(x => x.RefStudentID == e.Data.ID))
                    {
                        if (rewardInfo.SchoolYear == schoolYear && rewardInfo.Semester == semester && rewardInfo.OccurDate == occurdate && rewardInfo.Reason == reason)
                            match++;
                    }
                    if (match > 1)
                    {
                        e.ErrorMessage = "系統發現此事由在同一天中存在兩筆重複資料，無法進行更新，建議您手動處裡此筆變更。";
                        pass = false;
                    }
                    if (match == 0)
                    {
                        isInsert = true;
                    }
                    else
                    {
                        isUpdata = true;
                    }
                    #endregion
                }
                if (chose2.Checked)
                {
                    #region 以次數為Key更新
                    int awardA = 0;
                    int awardB = 0;
                    int awardC = 0;
                    int faultA = 0;
                    int faultB = 0;
                    int faultC = 0;
                    #region 驗證必填欄位
                    if (e.Data["大功"] != "" && !int.TryParse(e.Data["大功"], out awardA))
                    {
                        e.ErrorFields.Add("大功", "必需輸入數字");
                        pass = false;
                    }
                    if (e.Data["小功"] != "" && !int.TryParse(e.Data["小功"], out awardB))
                    {
                        e.ErrorFields.Add("小功", "必需輸入數字");
                        pass = false;
                    }
                    if (e.Data["嘉獎"] != "" && !int.TryParse(e.Data["嘉獎"], out awardC))
                    {
                        e.ErrorFields.Add("嘉獎", "必需輸入數字");
                        pass = false;
                    }
                    if (e.Data["大過"] != "" && !int.TryParse(e.Data["大過"], out faultA))
                    {
                        e.ErrorFields.Add("大過", "必需輸入數字");
                        pass = false;
                    }
                    if (e.Data["小過"] != "" && !int.TryParse(e.Data["小過"], out faultB))
                    {
                        e.ErrorFields.Add("小過", "必需輸入數字");
                        pass = false;
                    }
                    if (e.Data["警告"] != "" && !int.TryParse(e.Data["警告"], out faultC))
                    {
                        e.ErrorFields.Add("警告", "必需輸入數字");
                        pass = false;
                    }
                    #endregion
                    if (!pass)
                    {
                        return;
                    }
                    int match = 0;
                    #region 檢查重複
                    foreach (JHDisciplineRecord rewardInfo in CacheDiscipline.Values.Where(x => x.RefStudentID == e.Data.ID))
                    {
                        int MeritA = rewardInfo.MeritA.HasValue ? rewardInfo.MeritA.Value : 0;
                        int MeritB = rewardInfo.MeritB.HasValue ? rewardInfo.MeritB.Value : 0;
                        int MeritC = rewardInfo.MeritC.HasValue ? rewardInfo.MeritC.Value : 0;

                        int DemeritA = rewardInfo.DemeritA.HasValue ? rewardInfo.DemeritA.Value : 0;
                        int DemeritB = rewardInfo.DemeritB.HasValue ? rewardInfo.DemeritB.Value : 0;
                        int DemeritC = rewardInfo.DemeritC.HasValue ? rewardInfo.DemeritC.Value : 0;

                        if (rewardInfo.SchoolYear == schoolYear &&
                            rewardInfo.Semester == semester &&
                            rewardInfo.OccurDate == occurdate &&
                            MeritA == awardA &&
                            MeritB == awardB &&
                            MeritC == awardC &&
                            DemeritA == faultA &&
                            DemeritB == faultB &&
                            DemeritC == faultC)
                            match++;
                    }
                    #endregion
                    if (match > 1)
                    {
                        e.ErrorMessage = "系統發現此獎懲次數在同一天中存在兩筆重複資料，無法進行更新，建議您手動處裡此筆變更。";
                        pass = false;
                    }
                    if (match == 0)
                    {
                        isInsert = true;
                    }
                    else
                    {
                        isUpdata = true;
                    }
                    #endregion
                }
                if (!pass)
                {
                    allPass = false;
                    return;
                }
                #region 驗證可選則欄位值
                int integer;
                DateTime dateTime;
                bool hasAward = false, hasFault = false, IsErrorAward = true;
                foreach (string field in e.SelectFields)
                {
                    switch (field)
                    {
                        #region field
                        case "大功":
                        case "小功":
                        case "嘉獎":
                            if (e.Data[field] != "")
                            {
                                if (!int.TryParse(e.Data[field], out integer))
                                {
                                    e.ErrorFields.Add(field, "必需輸入數字");
                                    pass = false;
                                }
                                else
                                {
                                    if (integer < 0)
                                    {
                                        e.ErrorFields.Add(field, "不可為負數");
                                        pass = false;

                                    }
                                    else
                                    {
                                        IsErrorAward = false;
                                        hasAward |= integer > 0;
                                    }
                                }
                            }
                            break;
                        case "大過":
                        case "小過":
                        case "警告":
                            if (e.Data[field] != "")
                            {
                                if (!int.TryParse(e.Data[field], out integer))
                                {
                                    e.ErrorFields.Add(field, "必需輸入數字");
                                    pass = false;
                                }
                                else
                                {
                                    if (integer < 0)
                                    {
                                        e.ErrorFields.Add(field, "不可為負數");
                                        pass = false;

                                    }
                                    else
                                    {
                                        IsErrorAward = false;
                                        hasFault |= integer > 0;
                                    }
                                }
                            }
                            break;
                        case "銷過日期":
                            if (e.Data[field] != "" && !DateTime.TryParse(e.Data[field], out dateTime))
                            {
                                e.ErrorFields.Add(field, "輸入格式為 西元年//月//日");
                                pass = false;
                            }
                            break;
                        case "是否銷過":
                        case "留校察看":
                            if (e.Data[field] != "" && e.Data[field] != "是" && e.Data[field] != "否")
                            {
                                e.ErrorFields.Add(field, "如果為是請填入\"是\"否則請保留空白或填入\"否\"");
                                pass = false;
                            }
                            break;
                        case "登錄日期":
                            if (e.Data[field] != "" && !DateTime.TryParse(e.Data[field], out dateTime))
                            {
                                e.ErrorFields.Add(field, "輸入格式為 西元年//月//日");
                                pass = false;
                            }
                            break;
                        #endregion
                    }
                }

                bool 檢查是否獎都是0 = false;
                bool 檢查是否懲都是0 = false;
                bool 檢查獎是空值 = false;
                bool 檢查懲是空值 = false;
                int 大功A = 0;
                int 小功A = 0;
                int 嘉獎A = 0;
                int 大過A = 0;
                int 小過A = 0;
                int 警告A = 0;
                if (e.SelectFields.Contains("大功") && e.SelectFields.Contains("小功") && e.SelectFields.Contains("嘉獎"))
                {
                    #region 獎
                    //,是否都是空值
                    if (!string.IsNullOrEmpty(e.Data["大功"]) || !string.IsNullOrEmpty(e.Data["小功"]) || !string.IsNullOrEmpty(e.Data["嘉獎"]))
                    {

                        int.TryParse(e.Data["大功"], out 大功A);
                        int.TryParse(e.Data["小功"], out 小功A);
                        int.TryParse(e.Data["嘉獎"], out 嘉獎A);

                        //相加是否為0
                        if (大功A == 0 && 小功A == 0 && 嘉獎A == 0)
                        {
                            檢查是否獎都是0 = true;
                        }
                    }
                    else
                    {
                        檢查獎是空值 = true;
                    }
                    #endregion
                }

                if (e.SelectFields.Contains("大過") && e.SelectFields.Contains("小過") && e.SelectFields.Contains("警告"))
                {
                    #region 懲
                    //是否都是空值
                    if (!string.IsNullOrEmpty(e.Data["大過"]) || !string.IsNullOrEmpty(e.Data["小過"]) || !string.IsNullOrEmpty(e.Data["警告"]))
                    {
                        int.TryParse(e.Data["大過"], out 大過A);
                        int.TryParse(e.Data["小過"], out 小過A);
                        int.TryParse(e.Data["警告"], out 警告A);

                        if (大過A == 0 && 小過A == 0 && 警告A == 0)
                        {
                            檢查是否懲都是0 = true;
                        }
                    }
                    else
                    {
                        檢查懲是空值 = true;
                    }
                    #endregion
                }

                if (檢查是否獎都是0 && 檢查是否懲都是0)
                {
                    e.ErrorMessage = "獎懲皆為0,系統無法判斷此類型資料。";
                    pass = false;
                }
                else if (檢查是否獎都是0 && 檢查懲是空值)
                {
                    e.ErrorMessage = "獎懲皆為0,系統無法判斷此類型資料。";
                    pass = false;
                }
                else if (檢查是否懲都是0 && 檢查獎是空值)
                {
                    e.ErrorMessage = "獎懲皆為0,系統無法判斷此類型資料。";
                    pass = false;
                }

                if (IsErrorAward)
                {
                    e.ErrorMessage = "獎勵與懲戒不可同時未輸入內容!!";
                    pass = false;
                }

                if (hasAward && hasFault)
                {
                    e.ErrorMessage = "系統愚昧無法理解同時記功又記過的情況。";
                    pass = false;
                }

                if (!pass && isInsert && (!e.SelectFields.Contains("留校察看") || e.Data["留校察看"] != "是") && (!hasFault && !hasAward))
                {
                    e.ErrorMessage = "無法新增沒有獎懲的記錄。";
                    pass = false;
                }
                if (pass && isInsert)
                    insertRecords++;
                if (pass && isUpdata)
                    updataRecords++;

                #endregion
                if (!pass)
                {
                    allPass = false;
                }
                #endregion
            };
            wizard.ValidateComplete += delegate
            {
                StringBuilder sb = new StringBuilder();
                if (allPass && insertRecords > 0)
                {
                    sb.AppendLine("新增" + insertRecords + "筆獎懲記錄");
                }

                if (allPass && updataRecords > 0)
                {
                    sb.AppendLine("更新" + updataRecords + "筆獎懲記錄");
                }

                if (sb.ToString() != "")
                {
                    sb.AppendLine("\n如與資料筆數不符請勿繼續。");
                    MsgBox.Show(sb.ToString(), "新增與更新獎懲", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                }

                insertRecords = 0;
                updataRecords = 0;
            };
            wizard.ImportComplete += (sender, e) => MessageBox.Show("匯入完成!");
            wizard.ImportPackage += delegate(object sender, SmartSchool.API.PlugIn.Import.ImportPackageEventArgs e)
            {
                bool hasUpdate = false, hasInsert = false;

                List<JHDisciplineRecord> updateDisciplines = new List<JHDisciplineRecord>();
                List<JHDisciplineRecord> insertDisciplines = new List<JHDisciplineRecord>();

                //2014/3/6日新增Log記錄
                StringBuilder Log_sb = new StringBuilder();
                if (chose1.Checked)
                {
                    Log_sb.AppendLine("以事由更新獎懲記錄：");
                    Log_sb.AppendLine("");
                }
                if (chose2.Checked)
                {
                    Log_sb.AppendLine("以次數更新獎懲記錄：");
                    Log_sb.AppendLine("");
                }

                foreach (RowData row in e.Items)
                {
                    int schoolYear = int.Parse(row["學年度"]);
                    int semester = int.Parse(row["學期"]);
                    DateTime occurdate = DateTime.Parse(row["日期"]);
                    if (chose1.Checked)
                    {
                        #region 以事由為Key更新
                        bool isAward;
                        int awardA = 0;
                        int awardB = 0;
                        int awardC = 0;
                        int faultA = 0;
                        int faultB = 0;
                        int faultC = 0;
                        string cleared = string.Empty;
                        DateTime? cleardate = null;
                        DateTime? registerdate = null;
                        string clearreason = "";
                        bool ultimateAdmonition = false;

                        if (row.ContainsKey("大功"))
                            awardA = (row["大功"] == "") ? 0 : int.Parse(row["大功"]);
                        if (row.ContainsKey("小功"))
                            awardB = (row["小功"] == "") ? 0 : int.Parse(row["小功"]);
                        if (row.ContainsKey("嘉獎"))
                            awardC = (row["嘉獎"] == "") ? 0 : int.Parse(row["嘉獎"]);
                        if (row.ContainsKey("大過"))
                            faultA = (row["大過"] == "") ? 0 : int.Parse(row["大過"]);
                        if (row.ContainsKey("小過"))
                            faultB = (row["小過"] == "") ? 0 : int.Parse(row["小過"]);
                        if (row.ContainsKey("警告"))
                            faultC = (row["警告"] == "") ? 0 : int.Parse(row["警告"]);

                        cleared = e.ImportFields.Contains("是否銷過") ? row["是否銷過"] : string.Empty;

                        if (e.ImportFields.Contains("銷過日期") && row["銷過日期"] != "")
                            cleardate = DateTime.Parse(row["銷過日期"]);
                        //else 
                        //    cleardate = DateTime.Now;

                        clearreason = e.ImportFields.Contains("銷過事由") ? row["銷過事由"] : "";

                        ultimateAdmonition = e.ImportFields.Contains("留校察看") ? ((row["留校察看"] == "是") ? true : false) : false;

                        if (e.ImportFields.Contains("登錄日期") && row["登錄日期"] != "")
                            registerdate = DateTime.Parse(row["登錄日期"]);
                        //else 
                        //    registerdate = DateTime.Now;                        

                        string reason = row["事由"];
                        bool match = false;
                        foreach (JHDisciplineRecord rewardInfo in CacheDiscipline.Values.Where(x => x.RefStudentID == row.ID))
                        {
                            if (rewardInfo.SchoolYear == schoolYear && rewardInfo.Semester == semester && rewardInfo.OccurDate == occurdate && rewardInfo.Reason.Equals(reason))
                            {
                                match = true;
                                #region 其他項目
                                cleared = e.ImportFields.Contains("是否銷過") ? row["是否銷過"] : string.Empty;

                                if (e.ImportFields.Contains("銷過日期"))
                                    cleardate = row["銷過日期"] != "" ? K12.Data.DateTimeHelper.Parse(row["銷過日期"]) : null;
                                else
                                    cleardate = rewardInfo.ClearDate;

                                if (e.ImportFields.Contains("登錄日期"))
                                    registerdate = row["登錄日期"] != "" ? K12.Data.DateTimeHelper.Parse(row["登錄日期"]) : null;
                                else
                                    registerdate = rewardInfo.RegisterDate;

                                clearreason = e.ImportFields.Contains("銷過事由") ?
                                    row["銷過事由"] :
                                    rewardInfo.ClearReason;

                                //ultimateAdmonition = e.ImportFields.Contains("留校察看") ?
                                //    ((row["留校察看"] == "是") ? true : false) :
                                //    rewardInfo.UltimateAdmonition;

                                #endregion
                                JHDisciplineRecord record = new JHDisciplineRecord();

                                isAward = awardA + awardB + awardC > 0;

                                if (isAward)
                                {
                                    record.MeritA = awardA;
                                    record.MeritB = awardB;
                                    record.MeritC = awardC;
                                }
                                else
                                {
                                    record.DemeritA = faultA;
                                    record.DemeritB = faultB;
                                    record.DemeritC = faultC;
                                    record.Cleared = cleared;
                                    record.ClearDate = cleardate;
                                    record.Reason = clearreason;
                                }

                                record.MeritFlag = ultimateAdmonition ? "2" : isAward ? "1" : "0";
                                record.RefStudentID = row.ID;
                                record.SchoolYear = schoolYear;
                                record.Semester = semester;
                                record.OccurDate = occurdate;
                                record.RegisterDate = registerdate;
                                record.Reason = reason;
                                record.ID = rewardInfo.ID;

                                updateDisciplines.Add(record);

                                hasUpdate = true;
                                break;
                            }
                        }
                        if (!match)
                        {

                            JHDisciplineRecord record = new JHDisciplineRecord();

                            isAward = awardA + awardB + awardC > 0;
                            if (isAward)
                            {
                                record.MeritA = awardA;
                                record.MeritB = awardB;
                                record.MeritC = awardC;
                            }
                            else
                            {
                                record.DemeritA = faultA;
                                record.DemeritB = faultB;
                                record.DemeritC = faultC;
                                record.Cleared = cleared;
                                record.ClearDate = cleardate;
                                record.Reason = clearreason;
                            }

                            record.MeritFlag = ultimateAdmonition ? "2" : isAward ? "1" : "0";
                            record.RefStudentID = row.ID;
                            record.SchoolYear = schoolYear;
                            record.Semester = semester;
                            record.OccurDate = occurdate;
                            record.Reason = reason;
                            record.RegisterDate = registerdate;

                            insertDisciplines.Add(record);

                            hasInsert = true;
                        }
                        #endregion
                    }
                    if (chose2.Checked)
                    {
                        #region 以次數為Key更新
                        bool isAward;
                        int awardA = 0;
                        int awardB = 0;
                        int awardC = 0;
                        int faultA = 0;
                        int faultB = 0;
                        int faultC = 0;
                        string cleared = string.Empty;
                        DateTime? cleardate = null;
                        DateTime? registerdate = null;
                        string clearreason = "";
                        bool ultimateAdmonition = false;
                        string reason = row.ContainsKey("事由") ? row["事由"] : "";

                        if (row.ContainsKey("大功"))
                            awardA = (row["大功"] == "") ? 0 : int.Parse(row["大功"]);
                        if (row.ContainsKey("小功"))
                            awardB = (row["小功"] == "") ? 0 : int.Parse(row["小功"]);
                        if (row.ContainsKey("嘉獎"))
                            awardC = (row["嘉獎"] == "") ? 0 : int.Parse(row["嘉獎"]);
                        if (row.ContainsKey("大過"))
                            faultA = (row["大過"] == "") ? 0 : int.Parse(row["大過"]);
                        if (row.ContainsKey("小過"))
                            faultB = (row["小過"] == "") ? 0 : int.Parse(row["小過"]);
                        if (row.ContainsKey("警告"))
                            faultC = (row["警告"] == "") ? 0 : int.Parse(row["警告"]);
                        cleared = e.ImportFields.Contains("是否銷過") ? row["是否銷過"] : string.Empty;

                        if (e.ImportFields.Contains("銷過日期") && row["銷過日期"] != "")
                            cleardate = K12.Data.DateTimeHelper.Parse(row["銷過日期"]);

                        if (e.ImportFields.Contains("登錄日期") && row["登錄日期"] != "")
                            registerdate = K12.Data.DateTimeHelper.Parse(row["登錄日期"]);

                        clearreason = e.ImportFields.Contains("銷過事由") ?
                            row["銷過事由"] : "";

                        ultimateAdmonition = e.ImportFields.Contains("留校察看") ?
                            ((row["留校察看"] == "是") ? true : false) : false;

                        bool match = false;
                        foreach (JHDisciplineRecord rewardInfo in CacheDiscipline.Values.Where(x => x.RefStudentID == row.ID))
                        {
                            int MeritA = rewardInfo.MeritA.HasValue ? rewardInfo.MeritA.Value : 0;
                            int MeritB = rewardInfo.MeritB.HasValue ? rewardInfo.MeritB.Value : 0;
                            int MeritC = rewardInfo.MeritC.HasValue ? rewardInfo.MeritC.Value : 0;

                            int DemeritA = rewardInfo.DemeritA.HasValue ? rewardInfo.DemeritA.Value : 0;
                            int DemeritB = rewardInfo.DemeritB.HasValue ? rewardInfo.DemeritB.Value : 0;
                            int DemeritC = rewardInfo.DemeritC.HasValue ? rewardInfo.DemeritC.Value : 0;

                            if (rewardInfo.SchoolYear == schoolYear &&
                                rewardInfo.Semester == semester &&
                                rewardInfo.OccurDate == occurdate &&
                                MeritA == awardA &&
                                MeritB == awardB &&
                                MeritC == awardC &&
                                DemeritA == faultA &&
                                DemeritB == faultB &&
                                DemeritC == faultC)
                            {
                                match = true;
                                #region 其他項目
                                reason = e.ImportFields.Contains("事由") ? row["事由"] : rewardInfo.Reason;

                                cleared = e.ImportFields.Contains("是否銷過") ? row["是否銷過"] : string.Empty;

                                if (e.ImportFields.Contains("銷過日期"))
                                    cleardate = row["銷過日期"] != "" ? K12.Data.DateTimeHelper.Parse(row["銷過日期"]) : null;
                                else
                                    cleardate = rewardInfo.ClearDate;

                                if (e.ImportFields.Contains("登錄日期"))
                                    registerdate = row["登錄日期"] != "" ? K12.Data.DateTimeHelper.Parse(row["登錄日期"]) : null;
                                else
                                    registerdate = rewardInfo.RegisterDate;

                                clearreason = e.ImportFields.Contains("銷過事由") ?
                                    row["銷過事由"] :
                                    rewardInfo.ClearReason;

                                //ultimateAdmonition = e.ImportFields.Contains("留校察看") ?
                                //    ((row["留校察看"] == "是") ? true : false) :
                                //    rewardInfo.UltimateAdmonition;
                                #endregion

                                JHDisciplineRecord record = new JHDisciplineRecord();

                                isAward = awardA + awardB + awardC > 0;
                                if (isAward)
                                {
                                    record.MeritA = awardA;
                                    record.MeritB = awardB;
                                    record.MeritC = awardC;
                                }
                                else
                                {
                                    record.DemeritA = faultA;
                                    record.DemeritB = faultB;
                                    record.DemeritC = faultC;
                                    record.Cleared = cleared;
                                    record.ClearDate = cleardate;
                                    record.Reason = clearreason;
                                }

                                record.MeritFlag = ultimateAdmonition ? "2" : isAward ? "1" : "0";
                                record.RefStudentID = row.ID;
                                record.SchoolYear = schoolYear;
                                record.Semester = semester;
                                record.OccurDate = occurdate;
                                record.RegisterDate = registerdate;
                                record.Reason = reason;
                                record.ID = rewardInfo.ID;

                                updateDisciplines.Add(record);

                                hasUpdate = true;
                                break;
                            }
                        }
                        if (!match)
                        {
                            JHDisciplineRecord record = new JHDisciplineRecord();

                            isAward = awardA + awardB + awardC > 0;

                            if (isAward)
                            {
                                record.MeritA = awardA;
                                record.MeritB = awardB;
                                record.MeritC = awardC;
                            }
                            else
                            {
                                record.DemeritA = faultA;
                                record.DemeritB = faultB;
                                record.DemeritC = faultC;
                                record.Cleared = cleared;
                                record.ClearDate = cleardate;
                                record.Reason = clearreason;
                            }

                            record.MeritFlag = ultimateAdmonition ? "2" : isAward ? "1" : "0";
                            record.RefStudentID = row.ID;
                            record.SchoolYear = schoolYear;
                            record.Semester = semester;
                            record.OccurDate = occurdate;
                            record.RegisterDate = registerdate;
                            record.Reason = reason;

                            insertDisciplines.Add(record);

                            hasInsert = true;
                        }
                        #endregion
                    }
                }

                if (hasInsert)
                {
                    JHDiscipline.Insert(insertDisciplines);
                    Dictionary<string, K12.Data.StudentRecord> StudentDic = GetStudent(insertDisciplines);
                    Log_sb.AppendLine("新增" + "「" + updateDisciplines.Count + "」筆資料");
                    foreach (JHDisciplineRecord record in insertDisciplines)
                    {
                        if (StudentDic.ContainsKey(record.RefStudentID))
                            Log_sb.AppendLine(GetLogContext(record, StudentDic[record.RefStudentID]));
                    }
                }

                if (hasUpdate)
                {
                    JHDiscipline.Update(updateDisciplines);
                    Dictionary<string, K12.Data.StudentRecord> StudentDic = GetStudent(updateDisciplines);
                    Log_sb.AppendLine("更新" + "「" + updateDisciplines.Count + "」筆資料");
                    foreach (JHDisciplineRecord record in updateDisciplines)
                    {
                        if (StudentDic.ContainsKey(record.RefStudentID))
                            Log_sb.AppendLine(GetLogContext(record, StudentDic[record.RefStudentID]));
                    }
                }
                if (hasUpdate || hasInsert)
                {
                    FISCA.LogAgent.ApplicationLog.Log("匯入獎懲記錄", "新增或更新", Log_sb.ToString());
                }
            };
        }
        private string GetLogContext(JHDisciplineRecord record, K12.Data.StudentRecord studentRecord)
        {
            StringBuilder sb = new StringBuilder();
            if (record.MeritFlag == "1")
                sb.AppendLine("獎勵：");
            else if (record.MeritFlag == "0")
                sb.AppendLine("懲戒：");
            sb.Append("班級「" + studentRecord.Name + "」座號「" + studentRecord.Name + "」姓名「" + studentRecord.Name + "」");
            sb.AppendLine("日期「" + record.OccurDate.ToShortDateString() + "」事由「" + record.Reason + "」");
            return sb.ToString();
        }

        /// <summary>
        /// 取得學生清單
        /// </summary>
        private Dictionary<string, K12.Data.StudentRecord> GetStudent(List<JHDisciplineRecord> updateDisciplines)
        {
            Dictionary<string, K12.Data.StudentRecord> StudentDic = new Dictionary<string, K12.Data.StudentRecord>();
            List<string> StudentIDList = new List<string>();
            foreach (JHDisciplineRecord dis in updateDisciplines)
            {
                if (!StudentIDList.Contains(dis.RefStudentID))
                    StudentIDList.Add(dis.RefStudentID);
            }
            List<K12.Data.StudentRecord> StudentRecordList = K12.Data.Student.SelectByIDs(StudentIDList);
            foreach (K12.Data.StudentRecord stud in StudentRecordList)
            {
                if (!StudentDic.ContainsKey(stud.ID))
                {
                    StudentDic.Add(stud.ID, stud);
                }
            }
            return StudentDic;
        }
    }
}