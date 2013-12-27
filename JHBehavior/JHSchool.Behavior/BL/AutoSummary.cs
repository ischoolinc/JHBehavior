using System.Collections.Generic;
using System.Linq;
using System.Xml;
using JHSchool.Data;

namespace JHSchool.Behavior.BusinessLogic
{
    /// <summary>
    /// 自動計算學期缺曠獎懲統計
    /// </summary>
    public class AutoSummary
    {
        /// <summary>
        /// 根據多筆學生編號及學年度學期取得學期缺曠獎懲結算記錄物件
        /// </summary>
        /// <param name="StudentIDs">學生編號列表，若傳入null則會自動取得所有學生編號列表。</param>
        /// <param name="SchoolYear">學年度，若傳入null則會根據缺曠獎懲及日常生活表現紀錄來決定。</param>
        /// <param name="Semester">學期，若傳入null則會根據缺曠獎懲及日常生活表現紀錄來決定。</param>
        /// <returns></returns>
        public static List<AutoSummaryRecord> Select(IEnumerable<string> StudentIDs, IEnumerable<SchoolYearSemester> SchoolYearSemesters)
        {
            //建立自動計算缺曠獎懲物件列表
            List<AutoSummaryRecord> summaryrecords = new List<AutoSummaryRecord>();

            //若學生編號為null，則取得所有學生編號
            List<string> MappingStudentIDs = StudentIDs == null ? JHStudent.SelectAll().Select(x => x.ID).ToList() : StudentIDs.ToList();

            //若學年度學期為null，學年度也為null
            List<int> MappingSchoolYears = SchoolYearSemesters == null ? null : SchoolYearSemesters.Select(x => x.SchoolYear).ToList();

            //若學年度學期為null，學期也為null
            List<int> MappingSemesters = SchoolYearSemesters == null ? null : SchoolYearSemesters.Select(x => x.Semester).ToList();

            //將學年度學期轉為List型態
            List<SchoolYearSemester> MappingSchoolYearSemesters = new List<SchoolYearSemester>();

            //取得日常生活表現物件列表，只根據學生編號取得
            List<JHMoralScoreRecord> records = JHMoralScore.Select(null, MappingStudentIDs, null, null);

            //取得獎勵列表，根據學生編號及學年度、學期取得
            List<JHMeritRecord> meritrecords = JHMerit.Select(MappingStudentIDs, null, null, null, null, MappingSchoolYears, MappingSemesters);

            //取得懲戒列表，根據學生編號及學年度、學期取得
            List<JHDemeritRecord> demeritrecords = JHDemerit.Select(MappingStudentIDs, null, null, null, null, MappingSchoolYears, MappingSemesters);

            //取得缺曠記錄列表，根據學生編號及學年度、學期取得
            List<JHAttendanceRecord> attendancerecords = JHAttendance.Select(MappingStudentIDs, null, null, null, MappingSchoolYears, MappingSemesters);

            List<JHPeriodMappingInfo> PeriodMappingInfos = JHPeriodMapping.SelectAll();

            if (SchoolYearSemesters != null)
                MappingSchoolYearSemesters = SchoolYearSemesters.ToList();
            else //根據缺曠獎懲明細來決定學年度及學期
            {
                foreach (JHMeritRecord record in meritrecords)
                {
                    List<SchoolYearSemester> CurSchoolYearSemester = MappingSchoolYearSemesters.Where(x => x.SchoolYear == record.SchoolYear && x.Semester == record.Semester).ToList();

                    if (!(CurSchoolYearSemester.Count > 0))
                    {
                        SchoolYearSemester NewSchoolYearSemester = new SchoolYearSemester(record.SchoolYear, record.Semester);
                        MappingSchoolYearSemesters.Add(NewSchoolYearSemester);
                    }
                }

                foreach (JHDemeritRecord record in demeritrecords)
                {
                    if (record.Cleared.Equals(string.Empty))
                    {
                        List<SchoolYearSemester> CurSchoolYearSemester = MappingSchoolYearSemesters.Where(x => x.SchoolYear == record.SchoolYear && x.Semester == record.Semester).ToList();

                        if (!(CurSchoolYearSemester.Count > 0))
                        {
                            SchoolYearSemester NewSchoolYearSemester = new SchoolYearSemester(record.SchoolYear, record.Semester);
                            MappingSchoolYearSemesters.Add(NewSchoolYearSemester);
                        }
                    }
                }

                foreach (JHAttendanceRecord record in attendancerecords)
                {
                    List<SchoolYearSemester> CurSchoolYearSemester = MappingSchoolYearSemesters.Where(x => x.SchoolYear == record.SchoolYear && x.Semester == record.Semester).ToList();

                    if (!(CurSchoolYearSemester.Count > 0))
                    {
                        SchoolYearSemester NewSchoolYearSemester = new SchoolYearSemester(record.SchoolYear, record.Semester);
                        MappingSchoolYearSemesters.Add(NewSchoolYearSemester);
                    }
                }

                foreach (JHMoralScoreRecord record in records)
                {
                    List<SchoolYearSemester> CurSchoolYearSemester = MappingSchoolYearSemesters.Where(x => x.SchoolYear == record.SchoolYear && x.Semester == record.Semester).ToList();

                    if (!(CurSchoolYearSemester.Count > 0))
                    {
                        SchoolYearSemester NewSchoolYearSemester = new SchoolYearSemester(record.SchoolYear, record.Semester);
                        MappingSchoolYearSemesters.Add(NewSchoolYearSemester);
                    }
                }
            }

            foreach (string StudentID in StudentIDs)
                foreach (SchoolYearSemester SYS in MappingSchoolYearSemesters)
                {
                    AutoSummaryRecord summaryrecord = new AutoSummaryRecord();

                    summaryrecord.RefStudentID = StudentID;
                    summaryrecord.SchoolYear = SYS.SchoolYear;
                    summaryrecord.Semester = SYS.Semester;
                    summaryrecords.Add(summaryrecord);
                }

            foreach (AutoSummaryRecord summaryrecord in summaryrecords)
            {
                List<JHMoralScoreRecord> morals = records.Where(x => x.RefStudentID == summaryrecord.RefStudentID && x.SchoolYear == summaryrecord.SchoolYear && x.Semester == summaryrecord.Semester).ToList();

                if (morals.Count > 0)
                {
                    summaryrecord.AddInitialSummary(morals[0].InitialSummary);
                    summaryrecord.MoralScore = morals[0];
                }

                //學生清單
                List<JHStudentRecord> students = new List<JHStudentRecord>();

                students.Add(summaryrecord.Student);

                //根據學年度學期取得獎懲紀錄
                List<JHMeritRecord> merits = meritrecords.Where(x => x.RefStudentID == summaryrecord.RefStudentID && x.SchoolYear == summaryrecord.SchoolYear && x.Semester == summaryrecord.Semester).ToList();
                List<JHDemeritRecord> demerits = demeritrecords.Where(x => x.RefStudentID == summaryrecord.RefStudentID && x.SchoolYear == summaryrecord.SchoolYear && x.Semester == summaryrecord.Semester).ToList();

                summaryrecord.Merits = merits;
                summaryrecord.Demerits = demerits;

                int MeritA = 0;
                int MeritB = 0;
                int MeritC = 0;

                int DemeritA = 0;
                int DemeritB = 0;
                int DemeritC = 0;

                int ClearedDemeritA = 0;
                int ClearedDemeritB = 0;
                int ClearedDemeritC = 0;

                foreach (JHMeritRecord merit in merits)
                {
                    MeritA += merit.MeritA.HasValue ? merit.MeritA.Value : 0;
                    MeritB += merit.MeritB.HasValue ? merit.MeritB.Value : 0;
                    MeritC += merit.MeritC.HasValue ? merit.MeritC.Value : 0;
                }                

                foreach (JHDemeritRecord demerit in demerits)
                {
                    if (demerit.Cleared.Equals("是"))
                    {
                        ClearedDemeritA += demerit.DemeritA.HasValue ? demerit.DemeritA.Value : 0;
                        ClearedDemeritB += demerit.DemeritB.HasValue ? demerit.DemeritB.Value : 0;
                        ClearedDemeritC += demerit.DemeritC.HasValue ? demerit.DemeritC.Value : 0;
                    }
                    else 
                    {
                        DemeritA += demerit.DemeritA.HasValue ? demerit.DemeritA.Value : 0;
                        DemeritB += demerit.DemeritB.HasValue ? demerit.DemeritB.Value : 0;
                        DemeritC += demerit.DemeritC.HasValue ? demerit.DemeritC.Value : 0;
                    }
                }

                //將銷過數統計加到AutoSummary當中

                summaryrecord.ClearedDemeritA = ClearedDemeritA;
                summaryrecord.ClearedDemeritB = ClearedDemeritB;
                summaryrecord.ClearedDemeritC = ClearedDemeritC;

                //將獎懲紀錄的統計加到AutoSummary當中

                XmlElement MeritElm = summaryrecord.AutoSummary.SelectSingleNode("DisciplineStatistics/Merit") as XmlElement;

                MeritElm.SetAttribute("A", "" + (K12.Data.Int.Parse(MeritElm.GetAttribute("A")) + MeritA));
                MeritElm.SetAttribute("B", "" + (K12.Data.Int.Parse(MeritElm.GetAttribute("B")) + MeritB));
                MeritElm.SetAttribute("C", "" + (K12.Data.Int.Parse(MeritElm.GetAttribute("C")) + MeritC));

                XmlElement DemeritElm = summaryrecord.AutoSummary.SelectSingleNode("DisciplineStatistics/Demerit") as XmlElement;

                DemeritElm.SetAttribute("A", "" + (K12.Data.Int.Parse(DemeritElm.GetAttribute("A")) + DemeritA));
                DemeritElm.SetAttribute("B", "" + (K12.Data.Int.Parse(DemeritElm.GetAttribute("B")) + DemeritB));
                DemeritElm.SetAttribute("C", "" + (K12.Data.Int.Parse(DemeritElm.GetAttribute("C")) + DemeritC));

                //將缺曠紀錄統計加到AutoSummary當中
                List<JHAttendanceRecord> attendances = attendancerecords.Where(x => x.RefStudentID == summaryrecord.RefStudentID && x.SchoolYear == summaryrecord.SchoolYear && x.Semester == summaryrecord.Semester).ToList();

                summaryrecord.Attendances = attendances;

                Dictionary<string, int> attendancestat = new Dictionary<string, int>();
                Dictionary<string, string> PeriodTypeKeys = new Dictionary<string, string>();
                Dictionary<string, string> AbsenceTypeKeys = new Dictionary<string, string>();

                foreach (JHAttendanceRecord attendance in attendances)
                {
                    foreach (K12.Data.AttendancePeriod period in attendance.PeriodDetail)
                    {
                        List<JHPeriodMappingInfo> Infos = PeriodMappingInfos.Where(x => x.Name == period.Period).ToList();
                        string PeriodType = "{未定義}";
                        if (Infos.Count > 0)
                            PeriodType = Infos[0].Type;

                        string Key = PeriodType + "^-^" + period.AbsenceType;

                        if (!attendancestat.ContainsKey(Key))
                        {
                            attendancestat.Add(Key, 0);
                            PeriodTypeKeys.Add(Key, PeriodType);
                            AbsenceTypeKeys.Add(Key, period.AbsenceType);
                        }

                        attendancestat[Key]++;
                    }
                }

                foreach (string Key in attendancestat.Keys)
                {
                    string AbsenceType = AbsenceTypeKeys[Key];
                    string PeriodType = PeriodTypeKeys[Key];

                    XmlElement Elm = summaryrecord.AutoSummary.SelectSingleNode("AttendanceStatistics/Absence[@Name='" + AbsenceType + "' and @PeriodType='" + PeriodType + "']") as XmlElement;

                    if (Elm != null)
                    {
                        int Count = K12.Data.Int.Parse(Elm.GetAttribute("Count")) + attendancestat[Key];
                        Elm.SetAttribute("Count", "" + Count);
                    }
                    else
                    {
                        XmlElement NewElm = summaryrecord.AutoSummary.OwnerDocument.CreateElement("Absence");
                        NewElm.SetAttribute("Name", AbsenceType);
                        NewElm.SetAttribute("PeriodType", PeriodType);
                        NewElm.SetAttribute("Count", "" + attendancestat[Key]);
                        summaryrecord.AutoSummary.SelectSingleNode("AttendanceStatistics").AppendChild(NewElm);
                    }
                }
            }

            return summaryrecords;
        }

        /// <summary>
        /// 根據缺曠明細計算出缺曠統計值
        /// </summary>
        /// <param name="attendances"></param>
        /// <returns></returns>
        public static List<AbsenceCountRecord> Calculate(IEnumerable<JHAttendanceRecord> attendances)
        {
            //建立新的缺曠統計紀錄列列表
            List<AbsenceCountRecord> records = new List<AbsenceCountRecord>();
            //取得節次設定列表
            List<JHPeriodMappingInfo> PeriodMappingInfos = JHPeriodMapping.SelectAll();
            //假別統計物件
            Dictionary<string, int> attendancestat = new Dictionary<string, int>();
            Dictionary<string, string> PeriodTypeKeys = new Dictionary<string, string>();
            Dictionary<string, string> AbsenceTypeKeys = new Dictionary<string, string>();

            foreach (JHAttendanceRecord attendance in attendances)
            {
                foreach (K12.Data.AttendancePeriod period in attendance.PeriodDetail)
                {
                    List<JHPeriodMappingInfo> Infos = PeriodMappingInfos.Where(x => x.Name == period.Period).ToList();
                    string PeriodType = "{未定義}";
                    if (Infos.Count > 0)
                        PeriodType = Infos[0].Type;

                    string Key = PeriodType + "^-^" + period.AbsenceType;

                    if (!attendancestat.ContainsKey(Key))
                    {
                        attendancestat.Add(Key, 0);
                        PeriodTypeKeys.Add(Key, PeriodType);
                        AbsenceTypeKeys.Add(Key, period.AbsenceType);
                    }

                    attendancestat[Key]++;
                }
            }

            foreach (string Key in attendancestat.Keys)
            {
                string AbsenceType = AbsenceTypeKeys[Key];
                string PeriodType = PeriodTypeKeys[Key];

                AbsenceCountRecord record = new AbsenceCountRecord();

                record.Name = AbsenceType;
                record.PeriodType = PeriodType;
                record.Count = attendancestat[Key];

                records.Add(record);
            }

            return records;
        }
    }
}