using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using Framework;
using FISCA.DSAUtil;
using FISCA.Authentication;

namespace JHSchool.Behavior.Feature
{
    [AutoRetryOnWebException()]
    public class QueryDemerit
    {
        /// <summary>
        /// 取得所有懲戒資料。
        /// </summary>
        /// <returns></returns>
        public static List<DemeritRecord> GetAllDemeritRecords()
        {
            //組出Query清單
            StringBuilder req = new StringBuilder("<Request><Field><All/></Field></Request>");
            //一個DemeritRecord類別的List
            List<DemeritRecord> result = new List<DemeritRecord>();

            foreach (XmlElement item in DSAServices.CallService("SmartSchool.Student.UpdateRecord.GetDetailList", new DSRequest(req.ToString())).GetContent().GetElements("UpdateRecord"))
            {
                result.Add(new DemeritRecord(item.GetAttribute("RefStudentID"), item));
            }
            return result;
        }

        public static List<DemeritRecord> GetDemeritRecords(params string[] primaryKeys)
        {
            return GetDemeritRecords((IEnumerable<string>)primaryKeys);
        }

        /// <summary>
        /// 取得異動記錄資料。
        /// </summary>
        /// <param name="primaryKeys">學生編號清單。</param>
        /// <returns></returns>
        public static List<DemeritRecord> GetDemeritRecords(IEnumerable<string> primaryKeys)
        {
            bool haskey = false;

            StringBuilder req = new StringBuilder("<SelectRequest><Field><All/></Field><Condition>");
            foreach (string key in primaryKeys)
            {
                req.Append("<RefStudentID>" + key + "</RefStudentID>");
                haskey = true;
            }
            req.Append("<Or><MeritFlag>0</MeritFlag><MeritFlag>2</MeritFlag></Or>");    //MeritFlag=0 銷過,  MeritFlag=2 記過 , MeritFlag=1 記功
            req.Append("</Condition><Order><RefStudentID /><OccurDate>desc</OccurDate></Order></SelectRequest>");

            List<DemeritRecord> result = new List<DemeritRecord>();

           
            if (haskey)
            {
                //Invoke DSA Services and parse the response doc into DemeritRecord objects.
                foreach (XmlElement item in DSAServices.CallService("SmartSchool.Student.Discipline.GetDiscipline", new DSRequest(req.ToString())).GetContent().GetElements("Discipline"))
                {
                    result.Add(new DemeritRecord(item.GetAttribute("RefStudentID"), item));
                }
            }
            return result;
        }
        /// <summary>
        /// 取得功過獎懲換算規則
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        public static DSResponse GetDemeritStatistic(DSRequest request)
        {
            return FISCA.Authentication.DSAServices.CallService("SmartSchool.Student.Discipline.GetDemeritStatistic", request);
        }
    }
}
