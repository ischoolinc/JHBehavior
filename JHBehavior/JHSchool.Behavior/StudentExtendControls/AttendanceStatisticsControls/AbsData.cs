using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace JHSchool.Behavior.StudentExtendControls.AttendanceStatisticsControls
{
    public class AbsData
    {
        public List<AbsItem> Items { get; set; }

        public AbsData()
        {
            Items = new List<AbsItem>();
        }

        public void ClearItem()
        {
            Items.Clear();
        }

        public AbsItem GetItem(string schoolyear, string semester, string periodType, string name)
        {
            foreach (AbsItem item in Items)
            {
                if (item.SchoolYear == schoolyear && item.Semester == semester && item.PeriodType == periodType && item.Name == name)
                    return item;
            }
            return null;
        }

        public void SetItem(AbsItem absenceItem)
        {
            AbsItem item = GetItem(absenceItem.SchoolYear, absenceItem.Semester, absenceItem.PeriodType, absenceItem.Name);

            if (item == null)
                Items.Add(absenceItem);
            else
                item.Count = absenceItem.Count;
        }

        public XmlElement GetSemesterElement(string schoolyear, string semester)
        {
            XmlDocument doc = new XmlDocument();
            XmlElement element = doc.CreateElement("AttendanceStatistics");
            
            foreach (AbsItem item in Items)
            {
                if (item.SchoolYear != schoolyear || item.Semester != semester) continue;

                XmlElement absenceElement = doc.CreateElement("Absence");
                absenceElement.SetAttribute("Name", item.Name);
                absenceElement.SetAttribute("PeriodType", item.PeriodType);
                absenceElement.SetAttribute("Count", item.Count.ToString());

                element.AppendChild(absenceElement);
            }
            return element;
        }
    }
}
