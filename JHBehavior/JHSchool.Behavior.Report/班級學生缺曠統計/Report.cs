using _71103_classTDK;

namespace JHSchool.Behavior.Report.班級學生缺曠統計
{
    internal class Report : IReport
    {
        #region IReport 成員

        public void Print()
        {
            AttendanceStatistics F = new AttendanceStatistics();
            F.ShowDialog();
        }

        #endregion
    }
}
