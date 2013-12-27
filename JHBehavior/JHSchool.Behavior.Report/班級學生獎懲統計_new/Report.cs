
namespace JHSchool.Behavior.Report.班級學生獎懲統計
{
    class Report
    {
        public void Print()
        {
            MeritDemeritStatistics MDS = new MeritDemeritStatistics();
            MDS.ShowDialog();
        }
    }
}
