using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using Framework;


namespace JHSchool.Behavior
{
    /// <summary>
    /// 此類別代表學生的某一筆懲戒記錄。
    /// </summary>
    public class DemeritRecord
    {
        /// <summary>
        /// Constructor 
        /// </summary>
        internal DemeritRecord(string refStudentID, XmlElement element)
        {
            XmlHelper helper = new XmlHelper(element);

            RefStudentID = refStudentID;
            ID = helper.GetString("@ID");                                               //ID
            SchoolYear = helper.GetString("SchoolYear");                                //學年度
            Semester = helper.GetString("Semester");                                    //學期
            GradeYear = helper.GetString("GradeYear");                                  //年級
            ClassName = helper.GetString("ClassName");                                  //班級名稱
            SeatNo = helper.GetString("SeatNo");                                        //
            Gender = helper.GetString("Gender");                                        //姓別
            OccurDate = helper.GetDateString("OccurDate");                              //懲戒日期
            Type = helper.GetString("Type");                                            //
            StudentNumber = helper.GetString("StudentNumber");                          //學生姓名
            Reason = helper.GetString("Reason");                                        //事由
            RegisterDate = helper.GetString("RegisterDate");                                        //登錄日期
            DemeritA = helper.GetElement("Detail/Discipline/Demerit").Attributes["A"].Value;                //大過
            DemeritB = helper.GetElement("Detail/Discipline/Demerit").Attributes["B"].Value;                //小過
            DemeritC = helper.GetElement("Detail/Discipline/Demerit").Attributes["C"].Value;                //警告

            if (helper.GetElement("Detail/Discipline/Demerit").Attributes["ClearDate"] != null)
                ClearDate = helper.GetElement("Detail/Discipline/Demerit").Attributes["ClearDate"].Value;       //銷過日期

            if (helper.GetElement("Detail/Discipline/Demerit").Attributes["ClearReason"] != null)
                ClearReason = helper.GetElement("Detail/Discipline/Demerit").Attributes["ClearReason"].Value;   //銷過事由

            Cleared = helper.GetElement("Detail/Discipline/Demerit").Attributes["Cleared"].Value;           //銷過
            MeritFlag = helper.GetString("MeritFlag");                                  //0是懲戒,1是獎勵,2是留察
            
        }

        #region ========= Properties ========

        internal string RefStudentID { get; private set; }
        public string ID { get; private set; }
        /// <summary>
        /// 學期
        /// </summary>
        public string Semester { get; private set; }
        /// <summary>
        /// 日期
        /// </summary>
        public string OccurDate { get; private set; }

        public string Type { get; private set; }
        public string StudentNumber { get; private set; }
        /// <summary>
        /// 學年度
        /// </summary>
        public string SchoolYear { get; private set; }
        /// <summary>
        /// 年級
        /// </summary>
        public string GradeYear { get; private set; }
        //public string OriginDepartment { get; private set; }
        /// <summary>
        /// 0是懲戒,1是獎勵,2是留察
        /// </summary>
        public string MeritFlag { get; private set; }
        public string Name { get; private set; }

        public string SeatNo { get; private set; }
        public string Gender { get; private set; }
        public string ADNumber { get; private set; }
        public string ClassName { get; private set; }
        /// <summary>
        /// 事由
        /// </summary>
        public string Reason { get; private set; }
        /// <summary>
        /// 大過數
        /// </summary>
        public string DemeritA { get; private set; }
        /// <summary>
        /// 小過數
        /// </summary>
        public string DemeritB { get; private set; }
        /// <summary>
        /// 警告數
        /// </summary>
        public string DemeritC { get; private set; }

        /// <summary>
        /// 銷過日期
        /// </summary>
        public string ClearDate { get; private set; }
        /// <summary>
        /// 銷過事由
        /// </summary>
        public string ClearReason { get; private set; }
        /// <summary>
        /// 是否銷過
        /// </summary>
        public string Cleared { get; private set; }

        /// <summary>
        /// 登錄日期
        /// </summary>
        public string RegisterDate { get; private set; }

        #endregion

        // element 的 tag 結構為：

        //<Discipline ID="1097220">
        //    <Semester>1</Semester>
        //    <OccurDate>2007/12/13</OccurDate>
        //    <Type>1</Type>
        //    <StudentNumber>514163</StudentNumber>
        //    <SchoolYear>96</SchoolYear>
        //    <GradeYear>2</GradeYear>
        //    <MeritFlag>2</MeritFlag>
        //    <Name>陳文淇1</Name>
        //    <Detail>
        //        <Discipline>
        //            <Demerit A="0" B="0" C="0" ClearDate="" ClearReason="" Cleared="" />
        //        </Discipline>
        //    </Detail>
        //    <SeatNo />
        //    <Gender>女</Gender>
        //    <RefStudentID>169968</RefStudentID>
        //    <ClassName>綜二義</ClassName>
        //    <Reason>長太醜, 該死</Reason>
        //</Discipline>
    }
}
