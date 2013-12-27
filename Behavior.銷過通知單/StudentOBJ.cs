using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using K12.Data;

namespace JHBehavior.銷過通知單
{
    class StudentOBJ
    {
        public StudentOBJ()
        {
            DemeritStringList = new List<ClearDemeritDetail>();
            DemeritA = 0;
            DemeritB = 0;
            DemeritC = 0;
            DemeritSum = 0;
            DemeritSchoolA = 0;
            DemeritSchoolB = 0;
            DemeritSchoolC = 0;
        }
        /// <summary>
        /// 學生物件
        /// </summary>
        public StudentRecord student { get; set; }

        /// <summary>
        /// 大過
        /// </summary>
        public int DemeritA { get; set; }
        /// <summary>
        /// 小過
        /// </summary>
        public int DemeritB { get; set; }
        /// <summary>
        /// 警告
        /// </summary>
        public int DemeritC { get; set; }

        /// <summary>
        /// 懲戒換算加總器
        /// </summary>
        public int DemeritSum { get; set; }

        //學期懲戒
        public int DemeritSchoolA { get; set; }
        public int DemeritSchoolB { get; set; }
        public int DemeritSchoolC { get; set; }

        /// <summary>
        /// 懲戒明細內容
        /// </summary>
        public List<ClearDemeritDetail> DemeritStringList { get; set; }

        public string TeacherName { get; set; }
        public string ClassName{get;set;}
        public string SeatNo { get; set; }
        public string StudentNumber { get; set; }

        //收件人地址
        public string address { get; set; }
        //郵遞區號
        public string ZipCode { get; set; }
        //郵遞區號第一碼
        public string ZipCode1 { get; set; }
        //郵遞區號第二碼
        public string ZipCode2 { get; set; }
        //郵遞區號第三碼
        public string ZipCode3 { get; set; }
        //郵遞區號第四碼
        public string ZipCode4 { get; set; }
        //郵遞區號第五碼
        public string ZipCode5 { get; set; }

        //監護人
        public string CustodianName { get; set; }
        //父親
        public string FatherName { get; set; }
        //母親
        public string MotherName { get; set; }
    }
}
