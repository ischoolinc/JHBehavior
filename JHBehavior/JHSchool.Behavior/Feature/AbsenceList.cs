using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Framework;
using FISCA.DSAUtil;
using FISCA.Authentication;

//namespace JHSchool.Behavior.StudentExtendControls.AttendanceItemControls
namespace JHSchool.Behavior.Feature
{
    [AutoRetryOnWebException()]
    public class QueryAbsenceMapping
    {
        private static string GET_ABSENCE_LIST = "SmartSchool.Others.GetAbsenceList";

        public QueryAbsenceMapping() { }

        public static List<AbsenceMappingInfo> Load()
        {
            StringBuilder req = new StringBuilder("<Request><Field><Content/><All/></Field></Request>");
            List<AbsenceMappingInfo> result = new List<AbsenceMappingInfo>();

            foreach (XmlElement item in DSAServices.CallService(GET_ABSENCE_LIST, new DSRequest(req.ToString())).GetContent().GetElements("Absence"))
            {
                result.Add(new AbsenceMappingInfo(item));
            }

            return result;
        }
    }

}
