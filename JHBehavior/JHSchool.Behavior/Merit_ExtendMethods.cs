using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JHSchool.Behavior
{
    public static class Merit_ExtendMethods
    {
        /// <summary>
        /// 把 GetMerits 方法動態加到 StudentRecord 物件上。
        /// </summary>
        /// <param name="studentRec"></param>
        /// <returns></returns>
        public static List<MeritRecord> GetMerits(this StudentRecord studentRec)
        {
            return Merit.Instance[studentRec.ID];
        }

        /// <summary>
        /// 把 SyncMeritCache 方法動態加到 List<StudentRecord> 物件上。
        /// </summary>
        public static void SyncMeritCache(this IEnumerable<StudentRecord> studentRecs)
        {
            List<string> primaryKeys = new List<string>();
            foreach (var item in studentRecs)
            {
                primaryKeys.Add(item.ID);
            }
            Merit.Instance.SyncDataBackground(primaryKeys);
        }
    }
}
