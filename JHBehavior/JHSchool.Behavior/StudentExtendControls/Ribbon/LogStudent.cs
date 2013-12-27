using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JHSchool.Behavior.StudentExtendControls.Ribbon
{
    class LogStudent
    {
        public string StudentID { get; set; }

        /// <summary>
        /// 日期 節次,缺曠別
        /// </summary>
        public Dictionary<string, Dictionary<string, string>> beforeData = new Dictionary<string, Dictionary<string, string>>();
        
        /// <summary>
        /// 日期 節次,缺曠別
        /// </summary>
        public Dictionary<string, Dictionary<string, string>> afterData = new Dictionary<string, Dictionary<string, string>>();
        
        /// <summary>
        /// 刪除資料
        /// </summary>
        public List<string> deleteData = new List<string>();

    }
}
