
namespace JHSchool.Behavior.BusinessLogic
{
    /// <summary>
    /// 代表學年度、學期
    /// </summary>
    public struct SchoolYearSemester
    {
        /// <summary>
        /// 學年度
        /// </summary>
        public int SchoolYear { get; set; }
        /// <summary>
        /// 學期
        /// </summary>
        public int Semester { get; set; }

        /// <summary>
        /// 建構式
        /// </summary>
        /// <param name="schoolYear">學年度</param>
        /// <param name="semester">學期</param>
        public SchoolYearSemester(int schoolYear, int semester)
            : this()
        {
            this.SchoolYear = schoolYear;
            this.Semester = semester;
        }
    }
}