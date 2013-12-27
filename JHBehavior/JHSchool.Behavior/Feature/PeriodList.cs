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
    public class QueryPeriodMapping
    {
        private static string GET_PERIOD_LIST = "SmartSchool.Others.GetPeriodList";

        public QueryPeriodMapping() { }

        /// <summary>
        /// 取得節次對照表清單
        /// </summary>
        /// <returns></returns>
        public static List<PeriodMappingInfo> Load()
        {
            StringBuilder req = new StringBuilder("<Request><Field><Content/><All/></Field></Request>");
            List<PeriodMappingInfo> result = new List<PeriodMappingInfo>();

            foreach (XmlElement item in DSAServices.CallService(GET_PERIOD_LIST, new DSRequest(req.ToString())).GetContent().GetElements("Period"))
            {
                result.Add(new PeriodMappingInfo(item));
            }

            result.Sort(new PeriodComparer());

            return result;
        }
    }

    class PeriodComparer : IComparer<PeriodMappingInfo>
    {
        #region IComparer<PeriodInfo> 成員

        public int Compare(PeriodMappingInfo x, PeriodMappingInfo y)
        {
            if (x.Sort == y.Sort) return 0;
            else if (x.Sort > y.Sort) return 1;
            else return -1;
        }

        #endregion
    }
}
