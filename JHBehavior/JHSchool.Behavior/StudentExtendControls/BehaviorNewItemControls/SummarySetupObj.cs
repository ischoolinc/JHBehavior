using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using FISCA.DSAUtil;

namespace JHSchool.Behavior.StudentExtendControls
{
    class SummarySetupObj
    {
        private XmlElement _xml;
        private XmlElement AttendanceXmlElement;

        public SummarySetupObj(XmlElement xml)
        {
            _xml = xml;

            DisciplineStatisticsSetup();

            AttendanceStatisticsSetup();
        }

        //由Xml建立DisciplineStatistics資料物件
        private void DisciplineStatisticsSetup()
        {
            if (_xml.SelectSingleNode("DisciplineStatistics") != null)
            {
                XmlElement MeritXml = (XmlElement)_xml.SelectSingleNode("DisciplineStatistics/Merit");
                if (MeritXml != null)
                {
                    MeritA = ParseInt(MeritXml.GetAttribute("A"));
                    MeritB = ParseInt(MeritXml.GetAttribute("B"));
                    MeritC = ParseInt(MeritXml.GetAttribute("C"));
                }
                else
                {
                    MeritA = 0;
                    MeritB = 0;
                    MeritC = 0;
                }
                XmlElement DemeritXml = (XmlElement)_xml.SelectSingleNode("DisciplineStatistics/Demerit");
                if (DemeritXml != null)
                {
                    DemeritA = ParseInt(DemeritXml.GetAttribute("A"));
                    DemeritB = ParseInt(DemeritXml.GetAttribute("B"));
                    DemeritC = ParseInt(DemeritXml.GetAttribute("C"));
                }
                else
                {
                    DemeritA = 0;
                    DemeritB = 0;
                    DemeritC = 0;
                }

            }
            else //空的就是空的
            {
                MeritA = 0;
                MeritB = 0;
                MeritC = 0;
                DemeritA = 0;
                DemeritB = 0;
                DemeritC = 0;
            }
        }

        //判斷字串
        private int ParseInt(string LoadString)
        {
            int nowINT;
            if (int.TryParse(LoadString, out nowINT))
            {
                return nowINT;
            }
            else
            {
                return 0;
            }
        }

        //由Xml建立AttendanceStatistics資料物件
        private void AttendanceStatisticsSetup()
        {
            if (_xml.SelectSingleNode("AttendanceStatistics") != null)
            {
                foreach (XmlElement each in _xml.SelectNodes("AttendanceStatistics/Absence"))
                {
                    AttendanceSetupObj obj = new AttendanceSetupObj(each);
                    AttendanceList.Add(obj);
                }
            }
        }

        public List<AttendanceSetupObj> AttendanceList = new List<AttendanceSetupObj>();
        public int MeritA { get; set; }
        public int MeritB { get; set; }
        public int MeritC { get; set; }
        public int DemeritA { get; set; }
        public int DemeritB { get; set; }
        public int DemeritC { get; set; }

        public XmlElement GetAllXmlElement()
        {
            DSXmlHelper dsx = new DSXmlHelper("InitialSummary");
            if (AttendanceList.Count != 0)
            {
                dsx.AddElement("AttendanceStatistics");
                foreach (AttendanceSetupObj each in AttendanceList)
                {
                    dsx.AddElement("AttendanceStatistics", "Absence");
                    dsx.SetAttribute("AttendanceStatistics/Absence", "Count", each.Count.ToString());
                    dsx.SetAttribute("AttendanceStatistics/Absence", "Name", each.Name);
                    dsx.SetAttribute("AttendanceStatistics/Absence", "PeriodType", each.PeriodType);
                }
            }

            bool checkMerit = MeritA + MeritB + MeritC != 0;
            bool checkDemerit = DemeritA + DemeritB + DemeritC != 0;

            if (checkMerit || checkDemerit)
            {
                dsx.AddElement("DisciplineStatistics");

                dsx.AddElement("DisciplineStatistics", "Merit");
                dsx.SetAttribute("DisciplineStatistics/Merit", "A", MeritA.ToString());
                dsx.SetAttribute("DisciplineStatistics/Merit", "B", MeritB.ToString());
                dsx.SetAttribute("DisciplineStatistics/Merit", "C", MeritC.ToString());

                dsx.AddElement("DisciplineStatistics", "Demerit");
                dsx.SetAttribute("DisciplineStatistics/Demerit", "A", DemeritA.ToString());
                dsx.SetAttribute("DisciplineStatistics/Demerit", "B", DemeritB.ToString());
                dsx.SetAttribute("DisciplineStatistics/Demerit", "C", DemeritC.ToString());

            }

            return dsx.BaseElement;
        }

    }
}
