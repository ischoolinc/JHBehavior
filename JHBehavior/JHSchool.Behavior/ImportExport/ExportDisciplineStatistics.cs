using System.Linq;
using System.Collections.Generic;
using JHSchool.Behavior.BusinessLogic;
using SmartSchool.API.PlugIn;
using JHSchool.Data;

namespace JHSchool.Behavior.ImportExport
{
    class ExportDisciplineStatistics : SmartSchool.API.PlugIn.Export.Exporter
    {
        //建構子
        public ExportDisciplineStatistics()
        {
            this.Image = null;
            this.Text = "匯出獎勵懲戒統計";
        }

        //覆寫
        public override void InitializeExport(SmartSchool.API.PlugIn.Export.ExportWizard wizard)
        {
            wizard.ExportableFields.AddRange("學年度", "學期", "大功", "小功", "嘉獎", "大過", "小過", "警告");

            wizard.ExportPackage += (sender, e) =>
            {
                //取得選取學生的缺曠記錄                
                List<AutoSummaryRecord> data = AutoSummary.Select(e.List, null);

                //var sortrecords = from record in records orderby record.RefStudentID,record.SchoolYear,record.Semester select record;

                //added by Cloud 2014/1/2
                List<AutoSummaryRecord> records = new List<AutoSummaryRecord>();
                foreach (AutoSummaryRecord record in data)
                {
                    if (record.DemeritA + record.DemeritB + record.DemeritC + record.MeritA + record.MeritB + record.MeritC > 0)
                    {
                        records.Add(record);
                    }
                }
                records.Sort(SortStudent);

                //尋訪每個缺曠記錄
                foreach (AutoSummaryRecord record in records)
                {
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
                                case "學年度":
                                    row.Add(field, "" + record.SchoolYear);
                                    break;
                                case "學期":
                                    row.Add(field, "" + record.Semester);
                                    break;
                                case "大功":
                                    row.Add(field, "" + record.MeritA);
                                    break;
                                case "小功":
                                    row.Add(field, "" + record.MeritB);
                                    break;
                                case "嘉獎":
                                    row.Add(field, "" + record.MeritC);
                                    break;
                                case "大過":
                                    row.Add(field, "" + record.DemeritA);
                                    break;
                                case "小過":
                                    row.Add(field, "" + record.DemeritB);
                                    break;
                                case "警告":
                                    row.Add(field, "" + record.DemeritC);
                                    break;
                            }
                        }
                    }
                    e.Items.Add(row);
                }
            };
        }

        //modified by Cloud 2014/1/2
        private int SortStudent(AutoSummaryRecord xx, AutoSummaryRecord yy)
        {
            JHStudentRecord x = xx.Student;
            JHStudentRecord y = yy.Student;
            string xx1 = x.Class != null ? x.Class.Name : "";
            xx1 += x.SeatNo.HasValue ? x.SeatNo.Value.ToString().PadLeft(3, '0') : "000";
            xx1 += x.Name.PadLeft(10, '0');
            xx1 += xx.SchoolYear.ToString().PadLeft(4, '0');
            xx1 += xx.Semester.ToString().PadLeft(2, '0');

            string yy1 = y.Class != null ? y.Class.Name : "";
            yy1 += y.SeatNo.HasValue ? y.SeatNo.Value.ToString().PadLeft(3, '0') : "000";
            yy1 += y.Name.PadLeft(10, '0');
            yy1 += yy.SchoolYear.ToString().PadLeft(4, '0');
            yy1 += yy.Semester.ToString().PadLeft(2, '0');

            return xx1.CompareTo(yy1);
        }
    }
}