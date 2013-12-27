using System.Collections.Generic;
using JHSchool.Data;
using SmartSchool.API.PlugIn;
using System;

namespace JHSchool.Behavior.ImportExport
{
    class ExportMerit : SmartSchool.API.PlugIn.Export.Exporter
    {
        public ExportMerit()
        {
            this.Image = null;
            this.Text = "匯出獎勵記錄";
        }

        public override void InitializeExport(SmartSchool.API.PlugIn.Export.ExportWizard wizard)
        {
            wizard.ExportableFields.AddRange("學年度", "學期", "日期", "大功", "小功", "嘉獎", "事由","登錄日期");

            wizard.ExportPackage += delegate(object sender, SmartSchool.API.PlugIn.Export.ExportPackageEventArgs e)
            {
                List<JHStudentRecord> students = JHStudent.SelectByIDs(e.List);

                #region 獎勵資料(DicMerit)
                Dictionary<string, List<JHMeritRecord>> DicMerit = new Dictionary<string, List<JHMeritRecord>>();

                List<JHMeritRecord> ListMerit = JHMerit.SelectByStudentIDs(e.List);

                //ListMerit.Sort(SortDate);

                foreach (JHMeritRecord merit in ListMerit)
                {
                    if (!DicMerit.ContainsKey(merit.RefStudentID))
                    {
                        DicMerit.Add(merit.RefStudentID, new List<JHMeritRecord>());
                    }
                    DicMerit[merit.RefStudentID].Add(merit);
                }
                #endregion

                students.Sort(SortStudent);

                foreach (JHStudentRecord stud in students)
                {
                    if (DicMerit.ContainsKey(stud.ID))
                    {
                        DicMerit[stud.ID].Sort(SortDate);

                        foreach (JHMeritRecord JHR in DicMerit[stud.ID])
                        {
                            string OccurdateString = JHR.OccurDate.ToShortDateString();

                            string RegisterDateString = "";
                            if (JHR.RegisterDate.HasValue)
                            {
                                RegisterDateString = JHR.RegisterDate.Value.ToShortDateString();
                            }

                            RowData row = new RowData();
                            row.ID = stud.ID;
                            foreach (string field in e.ExportFields)
                            {
                                if (wizard.ExportableFields.Contains(field))
                                {
                                    switch (field)
                                    {
                                        case "學年度": row.Add(field, "" + JHR.SchoolYear.ToString()); break;
                                        case "學期": row.Add(field, "" + JHR.Semester.ToString()); break;
                                        case "日期": row.Add(field, "" + OccurdateString); break;
                                        case "大功": row.Add(field, "" + JHR.MeritA.ToString()); break;
                                        case "小功": row.Add(field, "" + JHR.MeritB.ToString()); break;
                                        case "嘉獎": row.Add(field, "" + JHR.MeritC.ToString()); break;
                                        case "事由": row.Add(field, "" + JHR.Reason); break;
                                        case "登錄日期": row.Add(field, "" + RegisterDateString); break;
                                    }
                                }
                            }
                            e.Items.Add(row);
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

        private int SortDate(JHMeritRecord x, JHMeritRecord y)
        {
            return x.OccurDate.CompareTo(y.OccurDate);
        }
    }
}
