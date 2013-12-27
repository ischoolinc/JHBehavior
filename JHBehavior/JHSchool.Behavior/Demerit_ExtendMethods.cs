using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JHSchool.Behavior
{
    public static class Demerit_ExtendMethods
    {
        /// <summary>
        /// 把 GetDemerits 方法動態加到 StudentRecord 物件上
        /// </summary>
        /// <param name="studentRec"></param>
        /// <returns></returns>
        public static List<DemeritRecord> GetDemerits(this StudentRecord studentRec)
        {
            return Demerit.Instance[studentRec.ID];
        }

        /// <summary>
        /// 把 SyncDemeritCache 方法動態加到 List<StudentRecord> 物件上
        /// </summary>
        /// <param name="studentRecs"></param>
        public static void SyncDemeritCache(this IEnumerable<StudentRecord> studentRecs)
        {
            List<string> primaryKeys = new List<string>();
            foreach (var item in studentRecs)
            {
                primaryKeys.Add(item.ID);
            }
            Demerit.Instance.SyncDataBackground(primaryKeys);
        }
    }
}
