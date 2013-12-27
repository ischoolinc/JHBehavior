using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

using Framework;
using FISCA.DSAUtil;
using FISCA.Authentication;

namespace JHSchool.Behavior
{
    /// <summary>
    /// Cache manager of Demerit
    /// </summary>
    public class Demerit : Framework.CacheManager<List<DemeritRecord>>
    {
        // Singleton pattern 
        private static Demerit _Instance = null;
        public static Demerit Instance { get { if (_Instance == null)_Instance = new Demerit(); return _Instance; } }
        private Demerit() { }


        /// <summary>
        /// 取得資料庫中所有的懲戒紀錄
        /// </summary>
        /// <returns></returns>
        protected override Dictionary<string, List<DemeritRecord>> GetAllData()
        {
            Dictionary<string, List<DemeritRecord>> oneToMany = new Dictionary<string, List<DemeritRecord>>();

            //foreach (DemeritRecord each in QueryDemerit.GetAllUpdateRecord())
            //{
            //    if (!oneToMany.ContainsKey(each.RefStudentID))
            //        oneToMany.Add(each.RefStudentID, new List<UpdateRecordRecord>());

            //    oneToMany[each.RefStudentID].Add(each);
            //}

            return oneToMany;
        }


        /// <summary>
        /// 取得指定的學生的懲戒紀錄
        /// </summary>
        /// <param name="primaryKeys">學生ID的集合</param>
        /// <returns></returns>
        protected override Dictionary<string, List<DemeritRecord>> GetData(IEnumerable<string> primaryKeys)
        {
            Dictionary<string, List<DemeritRecord>> result = new Dictionary<string, List<DemeritRecord>>();

            bool haskey = false;

            //建立 Request Document.
            StringBuilder req = new StringBuilder("<SelectRequest><Field><All/></Field><Condition>");
            foreach (string key in primaryKeys)
            {
                req.Append("<RefStudentID>" + key + "</RefStudentID>");
                haskey = true;
                result.Add(key, new List<DemeritRecord>());     //每一個傳入的 Key 都必須存在回傳的 Dictionary 中，否則不會觸發 ItemUpdated事件。
            }
            req.Append("<Or><MeritFlag>0</MeritFlag><MeritFlag>2</MeritFlag></Or>");    //MeritFlag=2 留校察看,  MeritFlag=0 記過 , MeritFlag=1 記功
            req.Append("</Condition><Order><RefStudentID /><OccurDate>desc</OccurDate></Order></SelectRequest>");

            //如果有傳學生ID進來
            if (haskey)
            {
                //Invoke DSA Services and parse the response doc into DemeritRecord objects.
                foreach (XmlElement item in DSAServices.CallService("SmartSchool.Student.Discipline.GetDiscipline", new DSRequest(req.ToString())).GetContent().GetElements("Discipline"))
                {
                    //item.GetAttribute("RefStudentID")
                    DemeritRecord dr = new DemeritRecord(item.SelectSingleNode("RefStudentID").InnerText, item);

                    //根據學生ID, 把 DemeritRecord 物件分配到相關的 List 中去。
                    if (!result.ContainsKey(dr.RefStudentID))
                        result.Add(dr.RefStudentID, new List<DemeritRecord>());

                    result[dr.RefStudentID].Add(dr);
                }
            }
            return result;            
        }
    }
}
