using System.Collections.Generic;
using JHSchool.Data;
using SmartSchool.API.PlugIn;
using System;
using K12.Data;

namespace JHSchool.Behavior.ImportExport
{
    class ExportDiscipline : SmartSchool.API.PlugIn.Export.Exporter
    {
        //建構子
        public ExportDiscipline()
        {
            this.Image = null;
            this.Text = "匯出獎懲紀錄";
        }

        public override void InitializeExport(SmartSchool.API.PlugIn.Export.ExportWizard wizard)
        {
            wizard.ExportableFields.AddRange("學年度", "學期", "日期", "大功", "小功", "嘉獎", "大過", "小過", "警告", "事由", "是否銷過", "銷過日期", "銷過事由", "登錄日期", "備註");

            wizard.ExportPackage += delegate (object sender, SmartSchool.API.PlugIn.Export.ExportPackageEventArgs e)
            {
                List<JHStudentRecord> students = JHStudent.SelectByIDs(e.List);

                #region 收集資料(DicMerit)
                Dictionary<string, List<DisciplineRecord>> DicDiscipline = new Dictionary<string, List<DisciplineRecord>>();

                List<DisciplineRecord> ListDiscipline = Discipline.SelectByStudentIDs(e.List);

                //ListDiscipline.Sort(SortDate);

                foreach (DisciplineRecord disRecord in ListDiscipline)
                {
                    if (!DicDiscipline.ContainsKey(disRecord.RefStudentID))
                    {
                        DicDiscipline.Add(disRecord.RefStudentID, new List<DisciplineRecord>());
                    }
                    DicDiscipline[disRecord.RefStudentID].Add(disRecord);
                }
                #endregion

                students = SortClassIndex.JHSchoolData_JHStudentRecord(students);

                foreach (JHStudentRecord stud in students)
                {
                    if (DicDiscipline.ContainsKey(stud.ID))
                    {

                        DicDiscipline[stud.ID].Sort(SortDate);

                        foreach (DisciplineRecord JHR in DicDiscipline[stud.ID])
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
                                        case "大功": row.Add(field, "" + JHR.MeritA.ToString()); break;
                                        case "小功": row.Add(field, "" + JHR.MeritB.ToString()); break;
                                        case "嘉獎": row.Add(field, "" + JHR.MeritC.ToString()); break;
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

        private int SortDate(DisciplineRecord x, DisciplineRecord y)
        {
            return x.OccurDate.CompareTo(y.OccurDate);
        }
    }
}
