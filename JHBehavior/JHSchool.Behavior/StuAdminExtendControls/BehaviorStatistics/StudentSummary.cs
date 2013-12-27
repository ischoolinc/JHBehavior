using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using FISCA.DSAUtil;
using JHSchool.Data;

namespace JHSchool.Behavior.StuAdminExtendControls.BehaviorStatistics
{
    internal class StudentSummary
    {
        #region 學生資料容器,包含資料,計算統計,產生XML功能

        public List<Data.JHAttendanceRecord> Attendances { get; private set; }
        public List<Data.JHMeritRecord> Merits { get; private set; }
        public List<Data.JHDemeritRecord> Demerits { get; private set; }

        //public List<AbsenceItem> InitAbsenceSummary { get; private set; }
        //public DisciplineItem InitMeritSummary { get; private set; }
        //public DisciplineItem InitDemeritSummary { get; private set; }

        public List<AbsenceItem> AbsenceSummary { get; private set; }
        public DisciplineItem MeritSummary { get; private set; }
        public DisciplineItem DemeritSummary { get; private set; }

        public Data.JHMoralScoreRecord OriginRecord { get; set; }

        public bool ContainsDetail = false;
        public bool ContainsInitial = false;

        public StudentSummary()
        {
            #region 建構子
            OriginRecord = null;

            Attendances = new List<JHSchool.Data.JHAttendanceRecord>();
            Merits = new List<JHSchool.Data.JHMeritRecord>();
            Demerits = new List<JHSchool.Data.JHDemeritRecord>();

            //InitAbsenceSummary = new List<AbsenceItem>();
            //InitMeritSummary = new DisciplineItem();
            //InitDemeritSummary = new DisciplineItem();

            AbsenceSummary = new List<AbsenceItem>();
            MeritSummary = new DisciplineItem();
            DemeritSummary = new DisciplineItem();
            #endregion
        }

        public void Calculate(PeriodMap map)
        {
            #region 計算統計功能
            DSXmlHelper DSX = new DSXmlHelper("Summary");

            Dictionary<string, int> dicAttendance = new Dictionary<string, int>();
            //計算Attendances(缺曠)至AbsenceSummary
            foreach (Data.JHAttendanceRecord each in Attendances) //缺曠每一天
            {
                foreach (K12.Data.AttendancePeriod savePeriod in each.PeriodDetail) //缺曠每一節
                {
                    string type = map.GetPeriodType(savePeriod.Period); //取得節次類型

                    IncreaseAbsenceCount(savePeriod.AbsenceType, type); //傳入節次 & 類型並統計

                }
            }
            //計算Merits(獎勵)至MeritSummary
            foreach (Data.JHMeritRecord each in Merits)
            {
                if (each.MeritA.HasValue)
                {
                    MeritSummary.A += each.MeritA.Value;
                }
                if (each.MeritB.HasValue)
                {
                    MeritSummary.B += each.MeritB.Value;
                }
                if (each.MeritC.HasValue)
                {
                    MeritSummary.C += each.MeritC.Value;
                }
            }
            //計算Demerits(懲戒)至DemeritSummary
            foreach (Data.JHDemeritRecord each in Demerits)
            {
                if (each.DemeritA.HasValue)
                {
                    DemeritSummary.A += each.DemeritA.Value;
                }
                if (each.DemeritB.HasValue)
                {
                    DemeritSummary.B += each.DemeritB.Value;
                }
                if (each.DemeritC.HasValue)
                {
                    DemeritSummary.C += each.DemeritC.Value;
                }
            }
            #endregion
        }

        private void IncreaseAbsenceCount(string name, string periodtype)
        {
            #region 尋找指定的缺曠,並統計

            //在AbsenceSummary,尋找指定的缺曠,並統計，不存在則建立。
            AbsenceItem item = null;

            foreach (AbsenceItem each in AbsenceSummary) //如果容器內有內容(表示InitialSummary有內容)
            {
                if (each.Name == name && each.Type == periodtype) //比對名稱類型
                {
                    item = each; //取得此統計
                    break;
                }
            }

            if (item == null) //如果沒有內容,則新增一筆
            {
                item = new AbsenceItem();
                item.Name = name;
                item.Type = periodtype;
                AbsenceSummary.Add(item);
            }

            item.Count++;
            #endregion
        }

        public XmlElement ToXml()
        {
            #region 以統計資料產生Xml結構
            DSXmlHelper newHelper = new DSXmlHelper("Summary");
            newHelper.AddElement("AttendanceStatistics");
            //缺曠
            foreach (AbsenceItem each in AbsenceSummary)
            {
                newHelper.AddElement("AttendanceStatistics", "Absence");
                newHelper.SetAttribute("AttendanceStatistics/Absence", "Count", each.Count.ToString());
                newHelper.SetAttribute("AttendanceStatistics/Absence", "Name", each.Name);
                newHelper.SetAttribute("AttendanceStatistics/Absence", "PeriodType", each.Type);
            }
            //獎勵
            newHelper.AddElement("DisciplineStatistics");
            newHelper.AddElement("DisciplineStatistics", "Merit");
            newHelper.SetAttribute("DisciplineStatistics/Merit", "A", MeritSummary.A.ToString());
            newHelper.SetAttribute("DisciplineStatistics/Merit", "B", MeritSummary.B.ToString());
            newHelper.SetAttribute("DisciplineStatistics/Merit", "C", MeritSummary.C.ToString());
            //懲戒
            newHelper.AddElement("DisciplineStatistics", "Demerit");
            newHelper.SetAttribute("DisciplineStatistics/Demerit", "A", DemeritSummary.A.ToString());
            newHelper.SetAttribute("DisciplineStatistics/Demerit", "B", DemeritSummary.B.ToString());
            newHelper.SetAttribute("DisciplineStatistics/Demerit", "C", DemeritSummary.C.ToString());

            return newHelper.BaseElement; //return Xml結構 
            #endregion
        }
        #endregion
    }
}
