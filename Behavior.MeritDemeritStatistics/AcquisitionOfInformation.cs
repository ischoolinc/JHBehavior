using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using K12.Data;
using K12.Logic;

namespace Behavior.MeritDemeritStatistics
{
    class AcquisitionOfInformation
    {
        #region 屬性
        public List<string> StudentList { get; set; } //取得資料用的ID
        public List<StudentRecord> HasNotDividedTheSex { get; set; } //未分性別

        public Dictionary<string, StudentRecord> StudentDic { get; set; } //Catch學生資料

        /// <summary>
        /// 一年級男生
        /// </summary>
        public List<AutoSummaryRecord> FirstGradeMale { get; set; }
        /// <summary>
        /// 二年級男生
        /// </summary>
        public List<AutoSummaryRecord> SecondGradeMale { get; set; }
        /// <summary>
        /// 三年級男生
        /// </summary>
        public List<AutoSummaryRecord> ThirdGradeMale { get; set; }

        /// <summary>
        /// 一年級女生
        /// </summary>
        public List<AutoSummaryRecord> FirstGradeFemale { get; set; }
        /// <summary>
        /// 二年級女生
        /// </summary>
        public List<AutoSummaryRecord> SecondGradeFemale { get; set; }
        /// <summary>
        /// 三年級女生
        /// </summary>
        public List<AutoSummaryRecord> ThirdGradeFemale { get; set; }

        /// <summary>
        /// 男生
        /// </summary>
        public List<AutoSummaryRecord> Male
        {
            get
            {
                List<AutoSummaryRecord> list = new List<AutoSummaryRecord>();
                list.AddRange(FirstGradeMale);
                list.AddRange(SecondGradeMale);
                list.AddRange(ThirdGradeMale);
                return list;
            }
        }

        /// <summary>
        /// 女生
        /// </summary>
        public List<AutoSummaryRecord> Female
        {
            get 
            {
                List<AutoSummaryRecord> list = new List<AutoSummaryRecord>();
                list.AddRange(FirstGradeFemale);
                list.AddRange(SecondGradeFemale);
                list.AddRange(ThirdGradeFemale);
                return list;
            }
        }

        /// <summary>
        /// 全部
        /// </summary>
        public List<AutoSummaryRecord> ALL
        {
            get
            {
                List<AutoSummaryRecord> list = new List<AutoSummaryRecord>();
                list.AddRange(FirstGradeMale);
                list.AddRange(SecondGradeMale);
                list.AddRange(ThirdGradeMale);
                list.AddRange(FirstGradeFemale);
                list.AddRange(SecondGradeFemale);
                list.AddRange(ThirdGradeFemale);
                return list;
            }
        }

        int SchoolYear { get; set; }
        int Semester { get; set; }
        bool IsSchoolYear { get; set; } //以學年為單位進行統計
        #endregion

        /// <summary>
        /// 傳入學年度學期
        /// 取得該學年度學期資料
        /// 包含學生清單與自動統計
        /// 並可依學年為單位進行統計
        /// </summary>
        public AcquisitionOfInformation(int _SchoolYear, int _Semester, bool _IsSchoolYear)
        {
            SchoolYear = _SchoolYear;
            Semester = _Semester;
            IsSchoolYear = _IsSchoolYear;
            StudentList = new List<string>();
            HasNotDividedTheSex = new List<StudentRecord>();

            FirstGradeMale = new List<AutoSummaryRecord>();
            SecondGradeMale = new List<AutoSummaryRecord>();
            ThirdGradeMale = new List<AutoSummaryRecord>();
            FirstGradeFemale = new List<AutoSummaryRecord>();
            SecondGradeFemale = new List<AutoSummaryRecord>();
            ThirdGradeFemale = new List<AutoSummaryRecord>();

            StudentDic = new Dictionary<string, StudentRecord>();

            GetStudent(); //取得學生

            GetAutoSummary(); //取得AutoSummary資料
        }

        //取得非刪除狀態之學生
        private void GetStudent()
        {
            foreach (StudentRecord student in Student.SelectAll())
            {
                //判斷學生狀態
                if (student.Status == StudentRecord.StudentStatus.刪除)
                    continue;
                //建立有效學生ID清單
                if (!StudentList.Contains(student.ID))
                {
                    StudentList.Add(student.ID);
                }
                //Catch學生資料
                if (!StudentDic.ContainsKey(student.ID))
                {
                    StudentDic.Add(student.ID, student);
                }
            }            
        }

        //取得AutoSummary並依男女分類
        private void GetAutoSummary()
        {
            Dictionary<string, SemesterHistoryItem> Dic = new Dictionary<string, SemesterHistoryItem>();
            List<string> SemesterHistoryStudentID = new List<string>();
            //取得所有的學期歷程
            foreach (SemesterHistoryRecord each in SemesterHistory.SelectByStudentIDs(StudentList))
            {
                //每名學生之學期歷程(多筆)
                foreach (SemesterHistoryItem item in each.SemesterHistoryItems)
                {
                    //取得使用者設定之(學年度 / 學期)歷程資料
                    if (IsSchoolYear)
                    {
                        if (item.SchoolYear == SchoolYear)
                        {
                            //避免重覆
                            if (!Dic.ContainsKey(each.RefStudentID))
                            {
                                SemesterHistoryStudentID.Add(each.RefStudentID);
                                Dic.Add(each.RefStudentID, item);
                            }
                        }
                    }
                    else
                    {
                        if (item.SchoolYear == SchoolYear && item.Semester == Semester)
                        {
                            //避免重覆
                            if (!Dic.ContainsKey(each.RefStudentID))
                            {
                                SemesterHistoryStudentID.Add(each.RefStudentID);
                                Dic.Add(each.RefStudentID, item);
                            }
                        }
                    }

                }
            }

            List<SchoolYearSemester> busiList = new List<SchoolYearSemester>();
            if (IsSchoolYear)
            {
                SchoolYearSemester sys = new SchoolYearSemester(SchoolYear, 1);
                busiList.Add(sys);
                sys = new SchoolYearSemester(SchoolYear, 2);
                busiList.Add(sys);
            }
            else
            {
                SchoolYearSemester sys = new SchoolYearSemester(SchoolYear, Semester);
                busiList.Add(sys);
            }

            //取得有使用者選定之學期歷程之學生
            List<AutoSummaryRecord> AutoSummaryList = AutoSummary.Select(SemesterHistoryStudentID, busiList);

            foreach (AutoSummaryRecord auto in AutoSummaryList)
            {
                StudentRecord student = StudentDic[auto.RefStudentID];

                if (!Dic.ContainsKey(auto.RefStudentID))
                    continue;

                if (student.Gender == "男")
                {
                    #region 男
                    if (Dic[auto.RefStudentID].GradeYear == 1 || Dic[auto.RefStudentID].GradeYear == 7)
                    {
                        FirstGradeMale.Add(auto);
                    }
                    if (Dic[auto.RefStudentID].GradeYear == 2 || Dic[auto.RefStudentID].GradeYear == 8)
                    {
                        SecondGradeMale.Add(auto);
                    }
                    if (Dic[auto.RefStudentID].GradeYear == 3 || Dic[auto.RefStudentID].GradeYear == 9)
                    {
                        ThirdGradeMale.Add(auto);
                    }
                    #endregion
                }
                else if (student.Gender == "女")
                {
                    #region 女
                    if (Dic[auto.RefStudentID].GradeYear == 1 || Dic[auto.RefStudentID].GradeYear == 7)
                    {
                        FirstGradeFemale.Add(auto);
                    }
                    if (Dic[auto.RefStudentID].GradeYear == 2 || Dic[auto.RefStudentID].GradeYear == 8)
                    {
                        SecondGradeFemale.Add(auto);
                    }
                    if (Dic[auto.RefStudentID].GradeYear == 3 || Dic[auto.RefStudentID].GradeYear == 9)
                    {
                        ThirdGradeFemale.Add(auto);
                    }
                    #endregion
                }
                else //未分性別
                {
                    HasNotDividedTheSex.Add(student); //記錄下未分性別之人
                }
            }
        }
    }
}
