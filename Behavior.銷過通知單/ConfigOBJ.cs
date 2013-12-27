using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace JHBehavior.銷過通知單
{
    class ConfigOBJ
    {
        /// <summary>
        /// 開始時間
        /// </summary>
        public DateTime StartDate { get; set; }
        /// <summary>
        /// 結束時間
        /// </summary>
        public DateTime EndDate { get; set; }
        /// <summary>
        /// 沒有資料即不印
        /// </summary>
        public bool PrintHasRecordOnly { get; set; }
        /// <summary>
        /// 範本
        /// </summary>
        public MemoryStream Template { get; set; }
        /// <summary>
        /// 寄件人姓名
        /// </summary>
        public string ReceiveName { get; set; }
        /// <summary>
        /// 寄件人地址
        /// </summary>
        public string ReceiveAddress { get; set; }
        /// <summary>
        /// 銷過日期 發生日期 登錄日期
        /// </summary>
        public string IsConditions { get; set; }

        /// <summary>
        /// 是否列印學生清單
        /// </summary>
        public bool PrintStudentList { get; set; }
    }
}
