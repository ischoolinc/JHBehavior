using System.Collections.Generic;
using JHSchool.Data;
using SmartSchool.API.PlugIn;

namespace JHSchool.Behavior.ImportExport
{
    class ExportAttendance : SmartSchool.API.PlugIn.Export.Exporter
    {
        //建構子
        public ExportAttendance()
        {
            this.Image = null;
            this.Text = "匯出缺曠記錄";
        }

        //覆寫
        public override void InitializeExport(SmartSchool.API.PlugIn.Export.ExportWizard wizard)
        {
            wizard.ExportableFields.AddRange("學年度", "學期","日期","缺曠假別","缺曠節次");

            wizard.ExportPackage += delegate(object sender, SmartSchool.API.PlugIn.Export.ExportPackageEventArgs e)
            {
                //取得學生清單
                List<JHStudentRecord> students = JHStudent.SelectByIDs(e.List);
                //取得學生相關缺曠記錄
                Dictionary<string, List<JHAttendanceRecord>> DicAttendance = new Dictionary<string, List<JHAttendanceRecord>>();

                List<JHAttendanceRecord> ListAttendance = JHAttendance.SelectByStudentIDs(e.List);

                foreach (JHAttendanceRecord attend in ListAttendance)
                {
                    if (!DicAttendance.ContainsKey(attend.RefStudentID))
                    {
                        DicAttendance.Add(attend.RefStudentID, new List<JHAttendanceRecord>());
                    }
                    DicAttendance[attend.RefStudentID].Add(attend);
                }

                students.Sort(SortStudent);
                
                //整理填入資料
                foreach (JHStudentRecord stud in students) //每一位學生
                {
                    if (DicAttendance.ContainsKey(stud.ID))
                    {
                        DicAttendance[stud.ID].Sort(SortDate);

                        foreach (JHAttendanceRecord att in DicAttendance[stud.ID]) //每一天
                        {
                            string OccurdateString = att.OccurDate.ToShortDateString();

                            foreach (K12.Data.AttendancePeriod Peroid in att.PeriodDetail) //每一節次
                            {
                                RowData row = new RowData();
                                row.ID = stud.ID;
                                foreach (string field in e.ExportFields)
                                {
                                    if (wizard.ExportableFields.Contains(field))
                                    {
                                        switch (field)
                                        {
                                            case "學年度": row.Add(field, "" + att.SchoolYear); break;
                                            case "學期": row.Add(field, "" + att.Semester); break;
                                            case "日期": row.Add(field, "" + OccurdateString); break;
                                            case "缺曠假別": row.Add(field, "" + Peroid.AbsenceType); break;
                                            case "缺曠節次": row.Add(field, "" + Peroid.Period); break;
                                        }
                                    }
                                }
                                e.Items.Add(row);
                            }
                        }
                    }
                }
            };
        }

        private int SortStudent(JHStudentRecord x, JHStudentRecord y)
        {

            string xx1 = x.Class != null ? x.Class.Name : "";
            string xx2 = x.SeatNo.HasValue ? x.SeatNo.Value.ToString().PadLeft(3, '0') : "000";
            string xx3 = xx1 + xx2;

            string yy1 = y.Class != null ? y.Class.Name : "";
            string yy2 = y.SeatNo.HasValue ? y.SeatNo.Value.ToString().PadLeft(3, '0') : "000";
            string yy3 = yy1 + yy2;

            return xx3.CompareTo(yy3);
        }

        private int SortDate(JHAttendanceRecord x, JHAttendanceRecord y)
        {
            return x.OccurDate.CompareTo(y.OccurDate);
        }
    }
}
