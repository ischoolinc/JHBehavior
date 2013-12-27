using System.Xml;

namespace JHSchool.Behavior.BusinessLogic
{
    public class AbsenceCountRecord
    {
        /// <summary>
        /// 缺曠次數
        /// </summary>
        public int Count { get; set; }

        /// <summary>
        /// 缺曠節次類別
        /// </summary>
        public string PeriodType { get; set; }

        /// <summary>
        /// 缺曠名稱
        /// </summary>
        public string Name { get; set; }

        public AbsenceCountRecord()
        {

        }

        public AbsenceCountRecord(XmlElement data)
        {
            Load(data);
        }

        public void Load(XmlElement data)
        {
            PeriodType = data.GetAttribute("PeriodType");
            Name = data.GetAttribute("Name");
            Count = K12.Data.Int.Parse(data.GetAttribute("Count"));
        }
    }
}