using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace JHSchool.Behavior
{
    /// <summary>
    /// 假別對照表的一筆資料
    /// </summary>
    public class AbsenceMappingInfo
    {
        //<Absence Abbreviation="曠" HotKey="c" Name="曠課" Noabsence="False" />
        //<Absence Abbreviation="事" HotKey="a" Name="事假" Noabsence="False" />

        public AbsenceMappingInfo() { }

        public AbsenceMappingInfo(XmlElement node)
        {
            Name = node.Attributes["Name"].InnerText;
            Abbreviation = node.Attributes["Abbreviation"].InnerText;
            HotKey = node.Attributes["HotKey"].InnerText;

            bool noabsence;
            if (!bool.TryParse(node.GetAttribute("Noabsence"), out noabsence))
                Noabsence = false;
            else
                Noabsence = true;
        }

        public string Name { get; private set; }
        public string Abbreviation { get; private set; }
        public string HotKey { get; private set; }
        public bool Noabsence { get; private set; }
    }
}
