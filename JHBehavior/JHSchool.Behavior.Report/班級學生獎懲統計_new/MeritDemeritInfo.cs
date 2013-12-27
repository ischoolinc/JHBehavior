
namespace JHSchool.Behavior.Report.班級學生獎懲統計
{
    class MeritDemeritInfo
    {

        /// <summary>
        /// 以班級為單位的物件
        /// </summary>
        public MeritDemeritInfo()
        {
        }

        /// <summary>
        /// 導師姓名
        /// </summary>
        public string TeacherName { set; get; }

        /// <summary>
        /// 班級名稱
        /// </summary>
        public string ClassName { set; get; }

        /// <summary>
        /// 班級學生人數
        /// </summary>
        public int StudentCount { set; get; }

        /// <summary>
        /// 獎勵人數
        /// </summary>
        public int MeritStudentCount { set; get; }

        /// <summary>
        /// 懲戒人數
        /// </summary>
        public int DemeritStudentCount { set; get; }

        /// <summary>
        /// 大功
        /// </summary>
        public int MeritA { set; get; }

        /// <summary>
        /// 小功
        /// </summary>
        public int MeritB { set; get; }

        /// <summary>
        /// 嘉獎
        /// </summary>
        public int MeritC { set; get; }

        /// <summary>
        /// 大過
        /// </summary>
        public int DemeritA { set; get; }

        /// <summary>
        /// 小過
        /// </summary>
        public int DemeritB { set; get; }

        /// <summary>
        /// 警告
        /// </summary>
        public int DemeritC { set; get; }

        /// <summary>
        /// 大功人數
        /// </summary>
        public int MeritAStudentCount { set; get; }

        /// <summary>
        /// 小功人數
        /// </summary>
        public int MeritBStudentCount { set; get; }

        /// <summary>
        /// 嘉獎人數
        /// </summary>
        public int MeritCStudentCount { set; get; }

        /// <summary>
        /// 大過人數
        /// </summary>
        public int DemeritAStudentCount { set; get; }

        /// <summary>
        /// 小過人數
        /// </summary>
        public int DemeritBStudentCount { set; get; }

        /// <summary>
        /// 警告人數
        /// </summary>
        public int DemeritCStudentCount { set; get; }

    }
}
