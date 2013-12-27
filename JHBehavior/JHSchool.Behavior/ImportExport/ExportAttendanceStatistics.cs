using System.Collections.Generic;
using JHSchool.Behavior.BusinessLogic;
using SmartSchool.API.PlugIn;
using JHSchool.Data;
using System.Linq;

namespace JHSchool.Behavior.ImportExport
{
    class ExportAttendanceStatistics : SmartSchool.API.PlugIn.Export.Exporter
    {
        //取得系統假別對照表
        private List<string> absenceList;

        //取得系統節次型態列表
        private List<string> periodtypelist;


        //建構子
        public ExportAttendanceStatistics()
        {
            this.Image = null;
            this.Text = "匯出缺曠統計";
        }

        //覆寫
        public override void InitializeExport(SmartSchool.API.PlugIn.Export.ExportWizard wizard)
        {
            wizard.ExportableFields.AddRange("學年度", "學期","缺曠假別","節次類型","缺曠統計值");

            wizard.ExportPackage += delegate(object sender, SmartSchool.API.PlugIn.Export.ExportPackageEventArgs e)
            {
                 absenceList = JHAbsenceMapping.SelectAll().Select(x => x.Name).ToList();

                 periodtypelist = JHPeriodMapping.SelectAll().Select(x => x.Type).ToList();

                //取得選取學生的缺曠記錄
                List<AutoSummaryRecord> records = AutoSummary.Select(e.List,null);

                //var sortrecords = from record in records orderby record.RefStudentID,record.SchoolYear,record.Semester select record;

                records.Sort(SortStudent);

                //尋訪每個缺曠記錄
                foreach (AutoSummaryRecord record in records)
                {
                    //尋訪每個缺曠統計值
                    foreach (AbsenceCountRecord countrecord in record.AbsenceCounts)
                    {
                        if (periodtypelist.Contains(countrecord.PeriodType) && absenceList.Contains(countrecord.Name))
                        {
                            //新增匯出列
                            RowData row = new RowData();

                            //指定匯出列的學生編號
                            row.ID = record.RefStudentID;

                            //判斷匯出欄位
                            foreach (string field in e.ExportFields)
                            {
                                if (wizard.ExportableFields.Contains(field))
                                {
                                    switch (field)
                                    {
                                        case "學年度": row.Add(field, "" + record.SchoolYear); break;
                                        case "學期": row.Add(field, "" + record.Semester); break;
                                        case "缺曠假別": row.Add(field, "" + countrecord.Name); break;
                                        case "節次類型": row.Add(field, "" + countrecord.PeriodType); break;
                                        case "缺曠統計值": row.Add(field, "" + countrecord.Count); break;
                                    }
                                }
                            }
                            e.Items.Add(row);
                        }
                    }
                }
            };
        }

        private int SortStudent(AutoSummaryRecord xx, AutoSummaryRecord yy)
        {
            JHStudentRecord x = xx.Student;
            JHStudentRecord y = yy.Student;
            string xx1 = x.Class != null ? x.Class.Name : "";
            string xx2 = x.SeatNo.HasValue ? x.SeatNo.Value.ToString().PadLeft(3, '0') : "000";
            string xx3 = xx1 + xx2;

            string yy1 = y.Class != null ? y.Class.Name : "";
            string yy2 = y.SeatNo.HasValue ? y.SeatNo.Value.ToString().PadLeft(3, '0') : "000";
            string yy3 = yy1 + yy2;

            return xx3.CompareTo(yy3);
        }
    }
}