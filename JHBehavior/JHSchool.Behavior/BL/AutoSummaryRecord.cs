using System.Collections.Generic;
using System.Xml;
using JHSchool.Data;

namespace JHSchool.Behavior.BusinessLogic
{
    /// <summary>
    /// 自動缺曠統計記錄物件
    /// </summary>
    public class AutoSummaryRecord
    {
        /// <summary>
        /// 學生編號
        /// </summary>
        public string RefStudentID { get; set; }

        /// <summary>
        /// 所屬學生紀錄物件
        /// </summary>
        public JHStudentRecord Student
        {
            get
            {
                if (!string.IsNullOrEmpty(RefStudentID))
                    return JHStudent.SelectByID(RefStudentID);
                else
                    return null;
            }
        }

        /// <summary>
        /// 學年度
        /// </summary>
        public int SchoolYear { get; set; }

        /// <summary>
        /// 學期
        /// </summary>
        public int Semester { get; set; }

        /// <summary>
        /// 自動統計之大功數
        /// </summary>
        public int MeritA
        {
            get
            {
                if (AutoSummary != null && !string.IsNullOrEmpty(AutoSummary.InnerXml))
                    return K12.Data.Int.Parse((AutoSummary.SelectSingleNode("DisciplineStatistics/Merit") as XmlElement).GetAttribute("A"));
                return 0;
            }
        }

        /// <summary>
        /// 非明細統計之大功數
        /// </summary>
        public int InitialMeritA
        {
            get
            {
                if (InitialSummary != null && !string.IsNullOrEmpty(InitialSummary.InnerXml))
                    return K12.Data.Int.Parse((InitialSummary.SelectSingleNode("DisciplineStatistics/Merit") as XmlElement).GetAttribute("A"));
                return 0;
            }
        }

        /// <summary>
        /// 自動統計之小功數
        /// </summary>
        public int MeritB
        {
            get
            {
                if (AutoSummary != null && !string.IsNullOrEmpty(AutoSummary.InnerXml))
                    return K12.Data.Int.Parse((AutoSummary.SelectSingleNode("DisciplineStatistics/Merit") as XmlElement).GetAttribute("B"));
                return 0;
            }
        }

        /// <summary>
        /// 非明細統計之小功數
        /// </summary>
        public int InitialMeritB
        {
            get
            {
                if (InitialSummary != null && !string.IsNullOrEmpty(InitialSummary.InnerXml))
                    return K12.Data.Int.Parse((InitialSummary.SelectSingleNode("DisciplineStatistics/Merit") as XmlElement).GetAttribute("B"));
                return 0;
            }
        }

        /// <summary>
        /// 自動統計之嘉獎數
        /// </summary>
        public int MeritC
        {
            get
            {
                if (AutoSummary != null && !string.IsNullOrEmpty(AutoSummary.InnerXml))
                    return K12.Data.Int.Parse((AutoSummary.SelectSingleNode("DisciplineStatistics/Merit") as XmlElement).GetAttribute("C"));
                return 0;
            }
        }

        /// <summary>
        /// 非明細統計之嘉獎數
        /// </summary>
        public int InitialMeritC
        {
            get
            {
                if (InitialSummary != null && !string.IsNullOrEmpty(InitialSummary.InnerXml))
                    return K12.Data.Int.Parse((InitialSummary.SelectSingleNode("DisciplineStatistics/Merit") as XmlElement).GetAttribute("C"));
                return 0;
            }
        }

        /// <summary>
        /// 自動統計之大過數
        /// </summary>
        public int DemeritA
        {
            get
            {
                if (AutoSummary != null && !string.IsNullOrEmpty(AutoSummary.InnerXml))
                    return K12.Data.Int.Parse((AutoSummary.SelectSingleNode("DisciplineStatistics/Demerit") as XmlElement).GetAttribute("A"));
                return 0;
            }
        }

        /// <summary>
        /// 非明細統計之大過數
        /// </summary>
        public int InitialDemeritA
        {
            get
            {
                if (InitialSummary != null && !string.IsNullOrEmpty(InitialSummary.InnerXml))
                    return K12.Data.Int.Parse((InitialSummary.SelectSingleNode("DisciplineStatistics/Demerit") as XmlElement).GetAttribute("A"));
                return 0;
            }
        }

        /// <summary>
        /// 銷過大過數
        /// </summary>
        public int ClearedDemeritA { get; set; }

        /// <summary>
        /// 自動統計之小過數
        /// </summary>
        public int DemeritB
        {
            get
            {
                if (AutoSummary != null && !string.IsNullOrEmpty(AutoSummary.InnerXml))
                    return K12.Data.Int.Parse((AutoSummary.SelectSingleNode("DisciplineStatistics/Demerit") as XmlElement).GetAttribute("B"));
                return 0;
            }
        }

        /// <summary>
        /// 非明細統計之小過數
        /// </summary>
        public int InitialDemeritB
        {
            get
            {
                if (InitialSummary != null && !string.IsNullOrEmpty(InitialSummary.InnerXml))
                    return K12.Data.Int.Parse((InitialSummary.SelectSingleNode("DisciplineStatistics/Demerit") as XmlElement).GetAttribute("B"));
                return 0;
            }
        }

        /// <summary>
        /// 銷過小過數
        /// </summary>
        public int ClearedDemeritB { get; set; }

        /// <summary>
        /// 自動統計之警告數
        /// </summary>
        public int DemeritC
        {
            get
            {
                if (AutoSummary != null && !string.IsNullOrEmpty(AutoSummary.InnerXml))
                    return K12.Data.Int.Parse((AutoSummary.SelectSingleNode("DisciplineStatistics/Demerit") as XmlElement).GetAttribute("C"));
                return 0;
            }
        }

        /// <summary>
        /// 非明細統計之警告數
        /// </summary>
        public int InitialDemeritC
        {
            get
            {
                if (InitialSummary != null && !string.IsNullOrEmpty(InitialSummary.InnerXml))
                    return K12.Data.Int.Parse((InitialSummary.SelectSingleNode("DisciplineStatistics/Demerit") as XmlElement).GetAttribute("C"));
                return 0;
            }
        }

        /// <summary>
        /// 銷過警告數
        /// </summary>
        public int ClearedDemeritC{ get; set;}

        /// <summary>
        /// 自動缺曠統計記錄列表
        /// </summary>
        public List<AbsenceCountRecord> AbsenceCounts
        {
            get
            {
                List<AbsenceCountRecord> records = new List<AbsenceCountRecord>();

                if (AutoSummary != null)
                {
                    foreach (XmlNode node in AutoSummary.SelectNodes("AttendanceStatistics/Absence"))
                    {
                        AbsenceCountRecord record = new AbsenceCountRecord();

                        record.Load(node as XmlElement);

                        records.Add(record);
                    }
                }

                return records;
            }
        }

        /// <summary>
        /// 非明細缺曠統計記錄列表
        /// </summary>
        public List<AbsenceCountRecord> InitialAbsenceCounts
        {
            get
            {
                List<AbsenceCountRecord> records = new List<AbsenceCountRecord>();

                if (InitialSummary != null)
                {
                    foreach (XmlNode node in InitialSummary.SelectNodes("AttendanceStatistics/Absence"))
                    {
                        AbsenceCountRecord record = new AbsenceCountRecord();

                        record.Load(node as XmlElement);

                        records.Add(record);
                    }
                }

                return records;
            }
        }

        /// <summary>
        /// 缺曠明細
        /// </summary>
        public List<JHAttendanceRecord> Attendances { get; set; }

        /// <summary>
        /// 獎勵明細
        /// </summary>
        public List<JHMeritRecord> Merits { get; set; }

        /// <summary>
        /// 懲戒明細
        /// </summary>
        public List<JHDemeritRecord> Demerits { get; set; }

        /// <summary>
        /// 學期缺曠獎懲統計，為非明細缺曠獎懲統計再加上系統缺曠獎懲紀錄，此屬性為在取得值時自動計算。
        /// </summary>
        public XmlElement AutoSummary { get; set; }

        /// <summary>
        /// 非明細缺曠獎懲統計
        /// </summary>
        public XmlElement InitialSummary { get; set; }

        /// <summary>
        /// 日常生活表現記錄物件
        /// </summary>
        public JHMoralScoreRecord MoralScore { get; set; }

        /// <summary>
        /// 初始化AutoSummary結構
        /// </summary>
        public AutoSummaryRecord()
        {
            //初始化要傳回的XML結構
            XmlDocument xmldoc = new XmlDocument();
            xmldoc.LoadXml("<Summary><AttendanceStatistics/><DisciplineStatistics><Merit A='0' B='0' C='0'/><Demerit A='0' B='0' C='0'/></DisciplineStatistics></Summary>");
            AutoSummary = xmldoc.DocumentElement;
        }

        /// <summary>
        /// 將InitialSummary的值複製到AutoSummary當中
        /// </summary>
        public void AddInitialSummary(XmlElement InitialSummary)
        {
            //將Summary的內容初始化成InitialSummary的內容
            if (InitialSummary != null)
            {
                this.InitialSummary = InitialSummary;

                //取得InitialSummary的獎勵統計
                XmlElement InitialMeritElement = InitialSummary.SelectSingleNode("DisciplineStatistics/Merit") as XmlElement;

                //獎勵統計的節點不為null
                if (InitialMeritElement != null)
                {
                    //取得Summary的獎勵統計
                    XmlElement AutoMeritElement = AutoSummary.SelectSingleNode("DisciplineStatistics/Merit") as XmlElement;

                    //設定大功值
                    string InitialA = string.IsNullOrEmpty(InitialMeritElement.GetAttribute("A")) ? "0" : InitialMeritElement.GetAttribute("A");
                    AutoMeritElement.SetAttribute("A", InitialA);

                    //設定小功值
                    string InitialB = string.IsNullOrEmpty(InitialMeritElement.GetAttribute("B")) ? "0" : InitialMeritElement.GetAttribute("B");
                    AutoMeritElement.SetAttribute("B", InitialB);

                    //設定嘉獎值
                    string InitialC = string.IsNullOrEmpty(InitialMeritElement.GetAttribute("C")) ? "0" : InitialMeritElement.GetAttribute("C");
                    AutoMeritElement.SetAttribute("C", InitialC);
                }

                //取得InitialSummary的懲戒統計
                XmlElement InitialDemeritElement = InitialSummary.SelectSingleNode("DisciplineStatistics/Demerit") as XmlElement;

                //懲戒統計的節點不為null
                if (InitialDemeritElement != null)
                {
                    //取得Summary的懲戒統計
                    XmlElement AutoDemeritElement = AutoSummary.SelectSingleNode("DisciplineStatistics/Demerit") as XmlElement;

                    //設定大過值
                    string InitialA = string.IsNullOrEmpty(InitialDemeritElement.GetAttribute("A")) ? "0" : InitialDemeritElement.GetAttribute("A");
                    AutoDemeritElement.SetAttribute("A", InitialA);

                    //設定小過值
                    string InitialB = string.IsNullOrEmpty(InitialDemeritElement.GetAttribute("B")) ? "0" : InitialDemeritElement.GetAttribute("B");
                    AutoDemeritElement.SetAttribute("B", InitialB);

                    //設定警告值
                    string InitialC = string.IsNullOrEmpty(InitialDemeritElement.GetAttribute("C")) ? "0" : InitialDemeritElement.GetAttribute("C");
                    AutoDemeritElement.SetAttribute("C", InitialC);
                }

                //取得InitialSummary的缺曠統計
                XmlElement InitialAttendanceElement = InitialSummary.SelectSingleNode("AttendanceStatistics") as XmlElement;

                if (InitialAttendanceElement != null)
                {
                    //取得Summary的缺曠統計
                    XmlElement AutoAttendanceElement = AutoSummary.SelectSingleNode("AttendanceStatistics") as XmlElement;

                    //將InitialSummary的缺曠統計加入到Summary中的缺曠統計
                    foreach (XmlElement Elm in InitialAttendanceElement.SelectNodes("Absence"))
                    {
                        XmlElement NewElm = AutoSummary.OwnerDocument.CreateElement("Absence");
                        NewElm.SetAttribute("Count", Elm.GetAttribute("Count"));
                        NewElm.SetAttribute("Name", Elm.GetAttribute("Name"));
                        NewElm.SetAttribute("PeriodType", Elm.GetAttribute("PeriodType"));

                        AutoAttendanceElement.AppendChild(NewElm);
                    }
                }
            }
        }
    }
}