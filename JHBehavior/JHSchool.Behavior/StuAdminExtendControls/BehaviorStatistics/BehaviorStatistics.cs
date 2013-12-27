using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SmartSchool.Common;
using Framework;
using System.Xml;
using FISCA.DSAUtil;
using JHSchool.Data;
using FISCA.LogAgent;

namespace JHSchool.Behavior.StuAdminExtendControls.BehaviorStatistics
{
    public partial class BehaviorStatistics : FISCA.Presentation.Controls.BaseForm
    {
        List<string> SortList = new List<string>();
        List<Data.JHClassRecord> ClassList = new List<Data.JHClassRecord>();
        BackgroundWorker BgW = new BackgroundWorker();
        //List<string> a;

        int SchoolYear;
        int Semester;
        string SeleteGradeYear;

        private List<Data.JHStudentRecord> LogStudRecord = new List<JHStudentRecord>(); //Log使用

        /// <summary>
        /// 以年級為群組的學生Record
        /// </summary>
        Dictionary<string, List<JHStudentRecord>> DicGradeYearStudent = new Dictionary<string, List<JHStudentRecord>>();

        public BehaviorStatistics()
        {
            InitializeComponent();
        }

        private void BehaviorStatistics_Load(object sender, EventArgs e)
        {

            #region 初始化畫面

            #region 年級資料處理
            ClassList.Clear();
            SortList.Clear();

            Data.JHClass.SelectAll();

            List<JHStudentRecord> StudentList = JHStudent.SelectAll();

            foreach (JHStudentRecord each in StudentList)
            {
                if (each.Class != null)
                {
                    if (each.Class.GradeYear != null)
                    {
                        string GradeYearString = each.Class.GradeYear.ToString();

                        if (!SortList.Contains(GradeYearString)) //建立年級清單
                        {
                            SortList.Add(GradeYearString);
                        }

                        if (!DicGradeYearStudent.ContainsKey(GradeYearString))
                        {
                            DicGradeYearStudent.Add(GradeYearString, new List<JHStudentRecord>());
                        }

                        DicGradeYearStudent[GradeYearString].Add(each);
                    }
                }
            }

            #region 舊寫法註解
            //ClassList = Data.JHClass.SelectAll();
            //foreach (Data.JHClassRecord eachClass in ClassList)
            //{
            //    if (eachClass.GradeYear == null)
            //    {
            //        //if (!SortList.Contains("未分年級"))
            //        //{
            //        //    SortList.Add("未分年級");                        
            //        //}
            //        //else
            //        //{
            //        continue;
            //        //}                                        
            //    }
            //    if (!SortList.Contains(eachClass.GradeYear.ToString()))
            //    {
            //        SortList.Add(eachClass.GradeYear.ToString());
            //    }
            //} 
            #endregion

            SortList.Sort();
            foreach (string each in SortList)
            {
                cbGradeYear.Items.Add(each);
            }
            cbGradeYear.SelectedIndex = 0;
            #endregion

            #region 學年度學期
            string schoolYear = K12.Data.School.DefaultSchoolYear;
            cbSchoolYear.Text = schoolYear;
            cbSchoolYear.Items.Add((int.Parse(schoolYear) - 2).ToString());
            cbSchoolYear.Items.Add((int.Parse(schoolYear) - 1).ToString());
            cbSchoolYear.Items.Add((int.Parse(schoolYear)).ToString());
            cbSchoolYear.Items.Add((int.Parse(schoolYear) + 1).ToString());

            string semester = K12.Data.School.DefaultSemester;
            cbSemester.Text = semester;
            cbSemester.Items.Add("1");
            cbSemester.Items.Add("2");
            #endregion

            BgW.DoWork += new DoWorkEventHandler(BgW_DoWork);
            BgW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BgW_RunWorkerCompleted);

            //iPeriodType = GetPeriodTypeItems();

            //GetMeritTypes()

            #endregion
        }

        private void btnSTART_Click(object sender, EventArgs e)
        {
            #region 開始計算
            this.Enabled = false;
            btnSTART.Text = "計算中!";
            xCount = 0;
            timer1.Enabled = true;

  
            SchoolYear = int.Parse(cbSchoolYear.Text);
            Semester = int.Parse(cbSemester.Text);

            if (!SortList.Contains(cbGradeYear.Text)) //取得學生資料
            {
                FISCA.Presentation.Controls.MsgBox.Show("輸入之年級不存在於系統中,請重新輸入");
                return;
            }

            SeleteGradeYear = cbGradeYear.Text;

            List<Data.JHStudentRecord> StudRecord = new List<Data.JHStudentRecord>();

            //foreach (Data.JHClassRecord each in ClassList) //將學生依選擇的年級加入清單
            //{
            //    if (each.GradeYear == null)
            //    {
            //        continue;
            //    }
            //    if (SeleteGradeYear == each.GradeYear)
            //    {
            //        StudRecord.AddRange(each.Students);
            //    }
            //}

            foreach (string each in DicGradeYearStudent.Keys)
            {
                if (each == SeleteGradeYear)
                {
                    StudRecord = DicGradeYearStudent[each];
                    break;
                }
            }


            List<string> StudentKey = new List<string>();

            LogStudRecord.Clear(); //Log使用

            foreach (Data.JHStudentRecord each in StudRecord) //將學生轉為Key
            {
                LogStudRecord.Add(each);

                //StudentKey.Add(each.ID);
            }

            BgW.RunWorkerAsync(StudRecord);
            #endregion
        }

        void BgW_DoWork(object sender, DoWorkEventArgs e)
        {
            List<JHStudentRecord> StudentKey = (List<JHStudentRecord>)e.Argument; //取得計算目標

            #region 初始化學生統計資料類別
            Dictionary<string, StudentSummary> stusummaries = new Dictionary<string, StudentSummary>();

            List<string> StudentSetKey = new List<string>();

            foreach (JHStudentRecord each in StudentKey) //對每個學生建立計算容器(類別StudentSummary)
            {
                stusummaries.Add(each.ID, new StudentSummary());
                StudentSetKey.Add(each.ID);
            }
            #endregion

            #region 取得缺曠資料
            List<Data.JHAttendanceRecord> attendances = Data.JHAttendance.SelectBySchoolYearAndSemester(StudentKey, SchoolYear, Semester); //取得缺曠資料(內容為一對多)

            foreach (Data.JHAttendanceRecord each in attendances) //一名學生有多筆記錄
            {
                stusummaries[each.RefStudentID].Attendances.Add(each); //裝入該學生的類別容器中

                //*****************
                stusummaries[each.RefStudentID].ContainsDetail = true;
            }
            #endregion

            #region 取得獎勵資料

            List<Data.JHMeritRecord> merit = Data.JHMerit.SelectBySchoolYearAndSemester(StudentSetKey, SchoolYear, Semester);

            foreach (Data.JHMeritRecord each in merit)
            {
                stusummaries[each.RefStudentID].Merits.Add(each);

                //*****************
                stusummaries[each.RefStudentID].ContainsDetail = true;
            }
            #endregion

            #region 取得懲戒資料
            List<Data.JHDemeritRecord> demerit = Data.JHDemerit.SelectByStudentIDs(StudentSetKey);

            foreach (Data.JHDemeritRecord each in demerit)
            {
                if (each.SchoolYear == SchoolYear && each.Semester == Semester)
                {
                    if (each.Cleared == "是") //排除已銷過內容
                    {
                        //忽略..
                    }
                    else
                    {
                        stusummaries[each.RefStudentID].Demerits.Add(each);
                    }

                    //*****************
                    stusummaries[each.RefStudentID].ContainsDetail = true;
                }
            }
            #endregion

            #region 取得InitialSummary資料

            List<Data.JHMoralScoreRecord> JHMoralScoreRecord = Data.JHMoralScore.SelectByStudentIDs(StudentSetKey); //JHMoralScoreRecord資料為一對多

            foreach (Data.JHMoralScoreRecord each in JHMoralScoreRecord)
            {
                if (each.SchoolYear == SchoolYear && each.Semester == Semester) //以學年度/學期篩選
                {
                    stusummaries[each.RefStudentID].OriginRecord = each; //記錄此筆JHMoralScoreRecord
                    //stusummaries[each.RefStudentID].Id = each.RefStudentID;

                    if (each.InitialSummary == null) continue; //如果內容為null  ->   <InitialSummary></InitialSummary>
                    if (string.IsNullOrEmpty(each.InitialSummary.InnerXml)) continue;

                    //*****************
                    stusummaries[each.RefStudentID].ContainsInitial = true;

                    XmlElement MSRecordXml = each.InitialSummary;

                    //缺曠資料
                    if (MSRecordXml.SelectNodes("AttendanceStatistics") != null)
                    {
                        foreach (XmlNode AbsencNode in MSRecordXml.SelectNodes("AttendanceStatistics/Absence"))
                        {
                            AbsenceItem Abs = new AbsenceItem();
                            XmlElement absence = AbsencNode as XmlElement; //轉

                            Abs.Count = ParseToInt(absence.GetAttribute("Count"));
                            Abs.Name = absence.GetAttribute("Name");
                            Abs.Type = absence.GetAttribute("PeriodType");
                            stusummaries[each.RefStudentID].AbsenceSummary.Add(Abs);
                        }
                    }

                    if (MSRecordXml.SelectSingleNode("DisciplineStatistics/Merit") != null)
                    {
                        //獎勵
                        XmlNode Merit = MSRecordXml.SelectSingleNode("DisciplineStatistics/Merit");
                        XmlElement ElementMerit = Merit as XmlElement; //轉
                        stusummaries[each.RefStudentID].MeritSummary.A = ParseToInt(ElementMerit.GetAttribute("A"));
                        stusummaries[each.RefStudentID].MeritSummary.B = ParseToInt(ElementMerit.GetAttribute("B"));
                        stusummaries[each.RefStudentID].MeritSummary.C = ParseToInt(ElementMerit.GetAttribute("C"));
                    }

                    if (MSRecordXml.SelectSingleNode("DisciplineStatistics/Demerit") != null)
                    {
                        //懲戒
                        XmlNode DeMerit = MSRecordXml.SelectSingleNode("DisciplineStatistics/Demerit");
                        XmlElement ElementDeMerit = DeMerit as XmlElement; //轉
                        stusummaries[each.RefStudentID].DemeritSummary.A = ParseToInt(ElementDeMerit.GetAttribute("A"));
                        stusummaries[each.RefStudentID].DemeritSummary.B = ParseToInt(ElementDeMerit.GetAttribute("B"));
                        stusummaries[each.RefStudentID].DemeritSummary.C = ParseToInt(ElementDeMerit.GetAttribute("C"));
                    }
                }
            }

            #endregion

            PeriodMap mapping = GetPeriodTypeItems(); //取得節次對照表

            foreach (StudentSummary each in stusummaries.Values) //取得每一筆StudentSummary
                each.Calculate(mapping); //執行其內部的Calculate()方法,並傳入節次對照表


            #region 整理新增&修改清單
            List<Data.JHMoralScoreRecord> inserts = new List<Data.JHMoralScoreRecord>(); //新增清單
            List<Data.JHMoralScoreRecord> updates = new List<Data.JHMoralScoreRecord>(); //修改清單

            foreach (KeyValuePair<string, StudentSummary> each in stusummaries)
            {
                if (each.Value.OriginRecord == null) //如果沒有Xml結構,表示需new一個JHMoralScoreRecord
                {
                    if (each.Value.ContainsDetail || each.Value.ContainsInitial) //如果有明細 或 有Initial,就要列入新增
                    {
                        Data.JHMoralScoreRecord record = new JHSchool.Data.JHMoralScoreRecord();
                        record.RefStudentID = each.Key; //必填 學生ID(建立stusummaries物件的key就是學生ID)
                        record.SchoolYear = SchoolYear; //必填 學年度(預設就是使用者畫面上選的 學年度 內容)
                        record.Semester = Semester; //必填 學期(預設就是使用者畫面上選的 學期 內容)
                        record.Summary = each.Value.ToXml(); //將Xml內容填入JHMoralScoreRecord物件的Summary欄位
                        inserts.Add(record); //加入新增清單
                    }
                }
                else //如果有
                {

                    if (radioButton1.Checked) //不覆蓋計算
                    {
                        if (each.Value.ContainsDetail || each.Value.ContainsInitial) //如果明細 或 Initial 任意有一資料,則不修改
                        {
                            each.Value.OriginRecord.Summary = each.Value.ToXml(); //將Xml內容填入JHMoralScoreRecord物件的Summary欄位
                            updates.Add(each.Value.OriginRecord); //加入修改清單
                        }
                    }
                    else //強制覆蓋計算計算
                    {
                        each.Value.OriginRecord.Summary = each.Value.ToXml(); //將Xml內容填入JHMoralScoreRecord物件的Summary欄位
                        updates.Add(each.Value.OriginRecord); //加入修改清單
                    }
                }
            } 
            #endregion

            #region 呼叫新增&修改服務
            try
            {
                if (inserts.Count != 0)
                {
                    Data.JHMoralScore.Insert(inserts);
                }

                if (updates.Count != 0)
                {
                    Data.JHMoralScore.Update(updates);
                }
            }
            catch (Exception ex)
            {
                btnSTART.Text = "開始計算";
                timer1.Enabled = false;
                FISCA.Presentation.Controls.MsgBox.Show("計算失敗,請重新執行此功能\n" + ex.Message);
                this.Enabled = true;
                return;
            }
            #endregion
        }

        void BgW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            #region 計算完成

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("詳細資料：");
            foreach (JHStudentRecord each in LogStudRecord)
            {
                if (each.Class != null)
                {
                    sb.AppendLine("姓名「" + each.Name + "」班級「" + each.Class.Name + "」座號「" + each.SeatNo + "」學號「" + each.StudentNumber + "」");
                }
                else
                {
                    sb.AppendLine("姓名「" + each.Name + "」學號「" + each.StudentNumber + "」");
                }
            }

            ApplicationLog.Log("學務系統.缺曠獎懲資料統計", "缺曠獎懲資料統計", "已進行「" + cbGradeYear.Text + "」年級「" + cbSchoolYear.Text + "」學年度「" + cbSemester.Text + "」學期，缺曠獎懲資料統計。\n統計對象共「" + LogStudRecord.Count + "」名學生。" + sb.ToString());
            this.Enabled = true;
            timer1.Enabled = false;
            btnSTART.Text = "開始計算";
            FISCA.Presentation.MotherForm.SetStatusBarMessage("缺曠統計資料已計算完成!", 100);

            FISCA.Presentation.Controls.MsgBox.Show("計算完成");
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");

            #endregion
        }

        public static int ParseToInt(string value)
        {
            #region Parse為數字
            if (string.IsNullOrEmpty(value)) return 0;

            return int.Parse(value);
            #endregion
        }

        public static PeriodMap GetPeriodTypeItems()
        {
            #region 取得節次類型

            ConfigData cd = School.Configuration["節次對照表"]; //取得節次對照表

            Dictionary<string, string> mapping = new Dictionary<string, string>();

            foreach (XmlElement each in cd.PreviousData.SelectNodes("Period"))
                mapping.Add(each.GetAttribute("Name"), each.GetAttribute("Type"));

            PeriodMap map = new PeriodMap(mapping); //new一個自定義的對照表

            return map;


            //string targetService = "SmartSchool.Config.GetList";
            //List<string> list = new List<string>();
            //DSXmlHelper helper = new DSXmlHelper("GetListRequest");
            //helper.AddElement("Field");
            //helper.AddElement("Field", "Content", "");
            //helper.AddElement("Condition");
            //helper.AddElement("Condition", "Name", "節次對照表");
            //DSRequest req = new DSRequest(helper.BaseElement);
            //DSResponse rsp = FISCA.Authentication.DSAServices.CallService(targetService, req);
            //foreach (XmlElement element in rsp.GetContent().GetElements("List/Periods/Period"))
            //{
            //    string type = element.GetAttribute("Type");
            //    if (!list.Contains(type))
            //        list.Add(type);
            //}
            //return list; 

            #endregion
        }

        public static List<string> GetMeritTypes()
        {
            #region 取得獎懲清單
            string targetService = "SmartSchool.Config.GetList";

            List<string> list = new List<string>();

            DSXmlHelper helper = new DSXmlHelper("GetListRequest");
            helper.AddElement("Field");
            helper.AddElement("Field", "Content");
            helper.AddElement("Condition");
            helper.AddElement("Condition", "Name", "鍵盤化獎懲資料管理_獎懲熱鍵表");

            DSRequest req = new DSRequest(helper.BaseElement);
            DSResponse rsp = FISCA.Authentication.DSAServices.CallService(targetService, req);

            foreach (XmlElement element in rsp.GetContent().GetElements("List/Configurations/Configuration"))
            {
                list.Add(element.GetAttribute("Name"));
            }
            return list;
            #endregion
        }

        private void BTNColse_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cbGradeYear_TextChanged(object sender, EventArgs e)
        {
            #region 年級檢查資料
            if (!SortList.Contains(cbGradeYear.Text))
            {
                errorProvider1.SetError(cbGradeYear, "年級必須為系統中已存在之範圍");
            }
            else
            {
                errorProvider1.Clear();
            }
            #endregion
        }

        private void cbSchoolYear_TextChanged(object sender, EventArgs e)
        {
            #region 學年度檢查
            int check;
            if (!int.TryParse(cbSchoolYear.Text, out check))
            {
                errorProvider2.SetError(cbSchoolYear, "學年度必須為數字內容");
            }
            else
            {
                errorProvider2.Clear();
            }
            #endregion
        }

        private void cbSemester_TextChanged(object sender, EventArgs e)
        {
            #region 學期檢查
            if (cbSemester.Text == "1" || cbSemester.Text == "2")
            {
                errorProvider3.Clear();
            }
            else
            {
                errorProvider3.SetError(cbSemester, "學期必須為 1 或 2");
            }
            #endregion
        }

        int xCount = 0;

        private void timer1_Tick(object sender, EventArgs e)
        {
            xCount++;
            FISCA.Presentation.MotherForm.SetStatusBarMessage("缺曠統計資料計算中...", xCount);
        }
    }
}
