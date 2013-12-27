using System;
using System.Collections.Generic;

namespace JHSchool.Behavior.Report.班級學生獎懲統計
{
    class ObjConfig
    {
        /// <summary>
        /// 開始時間
        /// </summary>
        public DateTime StartDate { set; get; }
        /// <summary>
        /// 結束時間
        /// </summary>
        public DateTime EndDate { set; get; }
        /// <summary>
        /// 是否包含銷過記錄
        /// </summary>
        public bool Cleared { set; get; }

        /// <summary>
        /// 選擇的獎懲清單
        /// </summary>
        public List<string> SelectItems { set; get; }

        /// <summary>
        /// 獎懲發生日期或登錄日期
        /// </summary>
        public bool InsertOrSetup { set; get; }
    }
}
