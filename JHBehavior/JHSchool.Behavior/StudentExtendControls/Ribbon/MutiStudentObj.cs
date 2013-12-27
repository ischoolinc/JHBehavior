using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JHSchool.Data;

namespace JHSchool.Behavior.StudentExtendControls.Ribbon
{
    class MutiStudentObj
    {
        /// <summary>
        /// 學生ID
        /// </summary>
        public string StudentID { get; set; }

        /// <summary>
        /// 學生缺曠清單
        /// </summary>
        public List<JHAttendanceRecord> AttendList = new List<JHAttendanceRecord>();

        /// <summary>
        /// 更新清單
        /// </summary>
        public List<JHAttendanceRecord> UpDataList = new List<JHAttendanceRecord>();

        /// <summary>
        /// 新增清單
        /// </summary>
        public List<JHAttendanceRecord> InsertList = new List<JHAttendanceRecord>();




        /// <summary>
        /// 傳入學生ID建立物件
        /// </summary>
        /// <param name="ID"></param>
        public MutiStudentObj(string ID)
        {
            StudentID = ID;
        }




        public void SetupAttendance(DateTime dt)
        {
            bool CheckTime = false;
            foreach (JHAttendanceRecord each in AttendList)
            {
                if (each.OccurDate == dt)
                {
                    UpDataList.Add(each);
                    CheckTime = true;
                }
            }

            if (!CheckTime)
            {

            }

            //處理UpDataList & InsertList




        }



    }
}
