using System.Collections.Generic;
using JHSchool.Data;
using System.Linq;
using SmartSchool.API.PlugIn;
using System;

namespace JHSchool.Behavior.ImportExport
{
    class ExportDemerit : SmartSchool.API.PlugIn.Export.Exporter
    {
        public ExportDemerit()
        {
            this.Image = null;
            this.Text = "匯出懲戒記錄";
        }

        public override void InitializeExport(SmartSchool.API.PlugIn.Export.ExportWizard wizard)
        {
            wizard.ExportableFields.AddRange("學年度", "學期", "日期", "大過", "小過", "警告", "事由","是否銷過","銷過日期","銷過事由","登錄日期","備註");

            wizard.ExportPackage += delegate(object sender, SmartSchool.API.PlugIn.Export.ExportPackageEventArgs e)
            {
                List<JHStudentRecord> students = JHStudent.SelectByIDs(e.List);

                Dictionary<string, List<JHDemeritRecord>> DicDemerit = new Dictionary<string, List<JHDemeritRecord>>();

                List<JHDemeritRecord> ListDemerit = JHDemerit.SelectByStudentIDs(e.List);

                //ListDemerit.Sort(SortDate);

                foreach (JHDemeritRecord demerit in ListDemerit)
                {
                    if (!DicDemerit.ContainsKey(demerit.RefStudentID))
                    {
                        DicDemerit.Add(demerit.RefStudentID, new List<JHDemeritRecord>());
                    }
                    DicDemerit[demerit.RefStudentID].Add(demerit);
                }

                students.Sort(SortStudent);

                foreach (JHStudentRecord stud in students)
                {
                    if (DicDemerit.ContainsKey(stud.ID))
                    {
                        DicDemerit[stud.ID].Sort(SortDate);

                        foreach (JHDemeritRecord JHR in DicDemerit[stud.ID])
                        {
                            string OccurdateString = JHR.OccurDate.ToShortDateString();

                            string ClearDateString = "";
                            if (JHR.ClearDate.HasValue)
                            {
                                ClearDateString = JHR.ClearDate.Value.ToShortDateString();
                            }

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
                                        case "大過": row.Add(field, "" + JHR.DemeritA.ToString()); break;
                                        case "小過": row.Add(field, "" + JHR.DemeritB.ToString()); break;
                                        case "警告": row.Add(field, "" + JHR.DemeritC.ToString()); break;
                                        case "事由": row.Add(field, "" + JHR.Reason); break;
                                        case "是否銷過": row.Add(field, "" + JHR.Cleared); break;
                                        case "銷過日期": row.Add(field, "" + ClearDateString); break;
                                        case "銷過事由": row.Add(field, "" + JHR.ClearReason); break;
                                        case "登錄日期": row.Add(field, "" + RegisterDateString); break;
                                        case "備註": row.Add(field, "" + JHR.Remark); break;
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

        private int SortDate(JHDemeritRecord x, JHDemeritRecord y)
        {
            return x.OccurDate.CompareTo(y.OccurDate);
        }
    }
}
