using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using JHSchool.Data;
using SmartSchool.API.PlugIn;

namespace JHSchool.Behavior.ImportExport
{
    class ImportAttendance : SmartSchool.API.PlugIn.Import.Importer
    {
        List<string> periodList = new List<string>();
        List<string> absenceList = new List<string>();
        List<string> importStudents = new List<string>();
        List<string> Keys = new List<string>();
        public ImportAttendance()
        {
            this.Image = null;
            this.Text = "匯入缺曠記錄";            
        }

        public override void InitializeImport(SmartSchool.API.PlugIn.Import.ImportWizard wizard)
        {
            wizard.PackageLimit = 1000;
            //缺曠記錄可匯入的欄位
            wizard.ImportableFields.AddRange("學年度", "學期", "日期", "缺曠假別", "缺曠節次");
            //缺曠記錄必需要有的欄位
            wizard.RequiredFields.AddRange("學年度", "學期", "日期", "缺曠假別", "缺曠節次");
            //開始驗證時的事件
            wizard.ValidateStart += new System.EventHandler<SmartSchool.API.PlugIn.Import.ValidateStartEventArgs>(wizard_ValidateStart);
            //驗證每行資料的事件
            wizard.ValidateRow += new System.EventHandler<SmartSchool.API.PlugIn.Import.ValidateRowEventArgs>(wizard_ValidateRow);
            //實際匯入資料的事件
            wizard.ImportPackage += new System.EventHandler<SmartSchool.API.PlugIn.Import.ImportPackageEventArgs>(wizard_ImportPackage); 
            //匯入完成事件
            wizard.ImportComplete += (sender, v) => MessageBox.Show("匯入完成");
        }

        void wizard_ValidateStart(object sender, SmartSchool.API.PlugIn.Import.ValidateStartEventArgs e)
        {
            //取得系統假別及節次對照表
            periodList = JHPeriodMapping.SelectAll().Select(x=>x.Name).ToList();
            absenceList = JHAbsenceMapping.SelectAll().Select(x=>x.Name).ToList();
            //Keys.Clear();
        }

        void wizard_ValidateRow(object sender, SmartSchool.API.PlugIn.Import.ValidateRowEventArgs e)
        {
            //"學年度", "學期", "日期", "缺曠假別", "缺曠節次"
            #region 驗各欄位填寫格式
            int t;
            foreach ( string field in e.SelectFields )
            {
                string value = e.Data[field];
                switch ( field )
                {
                    default:
                        break;
                    case "學年度":
                        if ( value == "" || !int.TryParse(value, out t) )
                            e.ErrorFields.Add(field, "此欄為必填欄位，必須填入整數。");
                        break;
                    case "學期":
                        if ( value == "" || !int.TryParse(value, out t) )
                        {
                            e.ErrorFields.Add(field, "此欄為必填欄位，必須填入整數。");
                        }
                        else if ( t != 1 && t != 2 )
                        {
                            e.ErrorFields.Add(field, "必須填入1或2");
                        }
                        break;
                    case "日期":
                        DateTime date = DateTime.Now;
                        if ( value == "" || !DateTime.TryParse(value, out date) )
                            e.ErrorFields.Add(field, "此欄為必填欄位，\n請依照\"西元年/月/日\"格式輸入。");
                        break;
                    case "缺曠假別":
                        if ( value == "" )
                        {
                            e.WarningFields.Add(field, "將會消除學生此筆缺曠記錄。");
                        }
                        else if ( !absenceList.Contains(value) )
                        {
                            e.ErrorFields.Add(field, "輸入值" + value + "不在假別清單中。");
                        }
                        break;
                    case "缺曠節次":
                        if ( !periodList.Contains(value) )
                            e.ErrorFields.Add(field, "輸入值" + value + "不在節次清單中。");
                        break;
                }
            }
            #endregion        
            #region 驗證主鍵
            string errorMessage = string.Empty;
            string Key = e.Data.ID + "-" + e.Data["日期"] + "-" + e.Data["缺曠節次"];

            if (Keys.Contains(Key))
                errorMessage = "學生編號、日期及缺曠節次的組合不能重覆!";
            else
                Keys.Add(Key);

            e.ErrorMessage = errorMessage;
            #endregion
        }

        void wizard_ImportPackage(object sender, SmartSchool.API.PlugIn.Import.ImportPackageEventArgs e)
        {
            //根據學生編號、學年度、學期及
            List<string> keyList = new List<string>();
            Dictionary<string, int> schoolYearMapping = new Dictionary<string, int>();
            Dictionary<string, int> semesterMapping = new Dictionary<string, int>();
            Dictionary<string, DateTime> dateMapping = new Dictionary<string, DateTime>();
            Dictionary<string, string> studentIDMapping = new Dictionary<string, string>();
            Dictionary<string, List<RowData>> rowsMapping = new Dictionary<string, List<RowData>>();
            Dictionary<string, string> oldPeriodTypes = new Dictionary<string, string>();
            Dictionary<string, List<JHAttendanceRecord>> studentAttendanceInfo = new Dictionary<string, List<JHAttendanceRecord>>();

            //1010913 - 新增檢查匯入日期,是否存在於不同學年度學期
            //ID / 日期 / 學年度學期
            //一名學生於不同學年度學期,不會有相同日期之資料
            Dictionary<string, Dictionary<string, List<string>>> testDic = new Dictionary<string, Dictionary<string, List<string>>>();

            //掃描每行資料，定出資料的PrimaryKey，並且將PrimaryKey對應到的資料寫成Dictionary
            foreach (RowData Row in e.Items)
            {
                int schoolYear = int.Parse(Row["學年度"]);
                int semester = int.Parse(Row["學期"]);
                DateTime date = DateTime.Parse(Row["日期"]);
                string studentID = Row.ID;

                #region 1010913 - 新增檢查匯入日期,是否存在於不同學年度學期
                if (!testDic.ContainsKey(studentID))
                {
                    testDic.Add(studentID, new Dictionary<string, List<string>>());
                }
                if (!testDic[studentID].ContainsKey(date.ToShortDateString()))
                {
                    testDic[studentID].Add(date.ToShortDateString(), new List<string>());
                }
                if (!testDic[studentID][date.ToShortDateString()].Contains(schoolYear + "_" + semester))
                {
                    testDic[studentID][date.ToShortDateString()].Add(schoolYear + "_" + semester);
                }
                #endregion   


                //string key = schoolYear + "^_^" + semester + "^_^" + date.ToShortDateString() + "^_^" + studentID;
                string key = studentID + "^_^"+ date.ToShortDateString();

                if (!keyList.Contains(key))
                {
                    keyList.Add(key);
                    schoolYearMapping.Add(key, schoolYear);
                    semesterMapping.Add(key, semester);
                    dateMapping.Add(key, date);
                    studentIDMapping.Add(key, studentID);
                    rowsMapping.Add(key, new List<RowData>());
                }
                rowsMapping[key].Add(Row);
            }

            List<K12.Data.StudentRecord> StudentList = K12.Data.Student.SelectByIDs(studentIDMapping.Values); ;
            Dictionary<string, K12.Data.StudentRecord> Students = new Dictionary<string, K12.Data.StudentRecord>();
            foreach (K12.Data.StudentRecord stu in StudentList)
            {
                if (!Students.ContainsKey(stu.ID))
                    Students.Add(stu.ID, stu);
            }


            #region 抓學生現有的缺曠記錄
            foreach (JHAttendanceRecord  var in JHAttendance.SelectByStudentIDs(studentIDMapping.Values.Distinct()))
            {
                if (!studentAttendanceInfo.ContainsKey(var.RefStudentID))
                    studentAttendanceInfo.Add(var.RefStudentID, new List<JHAttendanceRecord>());
                studentAttendanceInfo[var.RefStudentID].Add(var);

                #region 新增檢查匯入日期,是否存在於不同學年度學期(2)
                int schoolYear = var.SchoolYear;
                int semester = var.Semester;
                DateTime date = var.OccurDate;
                string studentID = var.RefStudentID;

                if (!testDic.ContainsKey(studentID))
                {
                    testDic.Add(studentID, new Dictionary<string, List<string>>());
                }
                if (!testDic[studentID].ContainsKey(date.ToShortDateString()))
                {
                    testDic[studentID].Add(date.ToShortDateString(), new List<string>());
                }
                if (!testDic[studentID][date.ToShortDateString()].Contains(schoolYear + "_" + semester))
                {
                    testDic[studentID][date.ToShortDateString()].Add(schoolYear + "_" + semester);
                } 
                #endregion
            }
            #endregion

            #region 1000512 - 新增檢查匯入日期,是否存在於不同學年度學期
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.AppendLine("匯入失敗!!您的匯入資料有以下狀況!!");
            bool sbbool = false;
            foreach (string each1 in testDic.Keys)
            {
                Dictionary<string, List<string>> dic = testDic[each1];
                K12.Data.StudentRecord sr = Students[each1]; //取得學生Record
                foreach (string each2 in dic.Keys)
                {
                    List<string> list = dic[each2];
                    if (list.Count > 1)
                    {
                        sb.Append("學生:「" + sr.Name + "」");
                        sb.Append("於不同學期");
                        foreach (string each3 in list)
                        {
                            sb.Append("「" + each3 + "」");
                        }
                        sb.AppendLine("有相同日期「" + each2 + "」資料!!");
                        sbbool = true;
                    }
                }
            }
            if (sbbool)
            {
                System.Windows.Forms.MessageBox.Show(sb.ToString());
                return;
            }
            #endregion

            List<JHAttendanceRecord> InsertAttendances = new List<JHAttendanceRecord>();
            List<JHAttendanceRecord> UpdateAttendances = new List<JHAttendanceRecord>();

            foreach (string key in keyList)
            {
                //根據學生編號、學年度、學期及日期取得缺曠記錄
                List<JHAttendanceRecord> records = new List<JHAttendanceRecord>();

                if (studentAttendanceInfo.ContainsKey(studentIDMapping[key]))
                {
                    records = studentAttendanceInfo[studentIDMapping[key]].Where(x => x.OccurDate == dateMapping[key]).ToList();
//                    records = studentAttendanceInfo[studentIDMapping[key]].Where(x => x.SchoolYear == schoolYearMapping[key] && x.Semester == semesterMapping[key] && x.OccurDate == dateMapping[key]).ToList();
                }

                //根據鍵值取得匯入資料，該匯入資料應該是有相同的學生編號、學年度、學期及缺曠日期
                List<RowData> Rows = rowsMapping[key];

                //該筆缺曠記錄已存在系統中
                if (records.Count > 0)
                {
                    //根據學生編號、學年度、學期及日期取得的缺曠記錄應該只有一筆
                    JHAttendanceRecord AttendanceRec = records[0];

                    for (int i = 0; i < Rows.Count; i++)
                    {
                        //取得匯入資料的節次及假別
                        string Period = Rows[i]["缺曠節次"];
                        string Absence = Rows[i]["缺曠假別"];

                        bool IsExist = false;

                        //節次已經存在會更新假別
                        foreach (K12.Data.AttendancePeriod CurrentPeriod in AttendanceRec.PeriodDetail)
                        {
                            if (CurrentPeriod.Period.Equals(Period))
                            {
                                CurrentPeriod.AbsenceType = Absence;
                                IsExist = true;
                            }
                        }

                        //若是節次不存在會根據節次及假別新增缺曠明細
                        if (!IsExist)
                        {
                            K12.Data.AttendancePeriod NewPeriod = new K12.Data.AttendancePeriod();
                            NewPeriod.AbsenceType = Absence;
                            NewPeriod.Period = Period;
                            AttendanceRec.PeriodDetail.Add(NewPeriod);
                        }
                    }

                    UpdateAttendances.Add(AttendanceRec);
                }
                else //該筆缺曠記錄沒有存在系統中
                {
                    JHAttendanceRecord record = new JHAttendanceRecord();

                    record.SchoolYear = schoolYearMapping[key];
                    record.Semester = semesterMapping[key];
                    record.OccurDate = dateMapping[key];
                    record.RefStudentID = Rows[0].ID;

                    //將屬於同樣一筆的匯入資料都加入到同樣的缺曠記錄中的明細
                    foreach (RowData Row in Rows)
                    {
                        K12.Data.AttendancePeriod Period = new K12.Data.AttendancePeriod();

                        Period.Period = Row["缺曠節次"];
                        Period.AbsenceType = Row["缺曠假別"];

                        record.PeriodDetail.Add(Period);
                    }

                    InsertAttendances.Add(record);
                }
            }

            if (InsertAttendances.Count>0)
                JHAttendance.Insert(InsertAttendances);

            if (UpdateAttendances.Count>0)
                JHAttendance.Update(UpdateAttendances);
        }
    }
}