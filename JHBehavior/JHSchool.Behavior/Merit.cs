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
    public class Merit:Framework.CacheManager<List<MeritRecord>>
    {
        // Singleton pattern 
        private static Merit _Instance = null;

        public static Merit Instance { get { if (_Instance == null)_Instance = new Merit(); return _Instance; } }
        
        private Merit() { }


        /// <summary>
        /// 取得資料庫中所有的懲勵紀錄
        /// </summary>
        /// <returns></returns>
        /// 
        protected override Dictionary<string, List<MeritRecord>> GetAllData()
        {
            Dictionary<string, List<MeritRecord>> oneToMany = new Dictionary<string, List<MeritRecord>>();

            return oneToMany;
        }


        /// <summary>
        /// 取得指定的學生的懲戒紀錄
        /// </summary>
        /// <param name="primaryKeys">學生ID的集合</param>
        /// <returns></returns>
        protected override Dictionary<string, List<MeritRecord>> GetData(IEnumerable<string> primaryKeys)
        {
            Dictionary<string, List<MeritRecord>> result = new Dictionary<string, List<MeritRecord>>();

            bool haskey = false;

            //建立 Request Document.
            StringBuilder req = new StringBuilder("<SelectRequest><Field><All/></Field><Condition>");
            foreach (string key in primaryKeys)
            {
                req.Append("<RefStudentID>" + key + "</RefStudentID>");
                haskey = true;
                result.Add(key, new List<MeritRecord>());     //每一個傳入的 Key 都必須存在回傳的 Dictionary 中，否則不會觸發 ItemUpdated事件。
            }
            req.Append("<MeritFlag>1</MeritFlag>");    //MeritFlag=0 銷過,  MeritFlag=2 記過 , MeritFlag=1 記功
            req.Append("</Condition><Order><RefStudentID /><OccurDate>desc</OccurDate></Order></SelectRequest>");

            //如果有傳學生ID進來
            if (haskey)
            {
                //Invoke DSA Services and parse the response doc into DemeritRecord objects.
                foreach (XmlElement item in DSAServices.CallService("SmartSchool.Student.Discipline.GetDiscipline", new DSRequest(req.ToString())).GetContent().GetElements("Discipline"))
                {
                    MeritRecord dr = new MeritRecord(item.SelectSingleNode("RefStudentID").InnerText, item);

                    //根據學生ID, 把 MeritRecord 物件分配到相關的 List 中去。
                    if (!result.ContainsKey(dr.RefStudentID))
                        result.Add(dr.RefStudentID, new List<MeritRecord>());

                    result[dr.RefStudentID].Add(dr);
                }
            }
            return result;            
        }
    }
}