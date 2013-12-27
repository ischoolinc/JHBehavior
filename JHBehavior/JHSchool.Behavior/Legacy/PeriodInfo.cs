using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace JHSchool.Behavior.Legacy
{
    public class PeriodInfo
    {
        public PeriodInfo()
        {
        }

        public PeriodInfo(XmlElement element)
        {
            _name = element.GetAttribute("Name");
            _type = element.GetAttribute("Type");
            string s = element.GetAttribute("Sort");
            if (!int.TryParse(s, out _sort))
                _sort = int.MaxValue;
            _aggregated = element.GetAttribute("Aggregated");
        }

        private string _name;

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        private int _sort;

        public int Sort
        {
            get { return _sort; }
            set { _sort = value; }
        }

        private string _type;

        public string Type
        {
            get { return _type; }
            set { _type = value; }
        }

        private string _aggregated;
        public string Aggregated
        {
            get { return _aggregated; }
        }
    }
}