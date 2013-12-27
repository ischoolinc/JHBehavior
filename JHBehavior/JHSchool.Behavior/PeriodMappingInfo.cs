using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace JHSchool.Behavior
{
    /// <summary>
    /// 代表節次對照表的一個節次
    /// </summary>
    public class PeriodMappingInfo
    {
        //<Period Aggregated="0.5" Name="早自習" Sort="1" Type="集會" />
        //<Period Aggregated="0.8" Name="升旗" Sort="2" Type="集會" />
        public PeriodMappingInfo() { }

        public PeriodMappingInfo(XmlElement node)
        {
            Name = node.Attributes["Name"].InnerText;
            Type = node.Attributes["Type"].InnerText;

            int sort;
            if (!int.TryParse(node.Attributes["Sort"].InnerText, out sort))
                Sort = int.MaxValue;
            else
                Sort = sort;

            float aggregated;
            if (!float.TryParse(node.GetAttribute("Aggregated"), out aggregated))
                Aggregated = 0.0f;
            else
                Aggregated = aggregated;
        }

        public string Name { get; private set; }
        public string Type { get; private set; }
        public int Sort { get; private set; }
        public float Aggregated { get; private set; }
    }
}
