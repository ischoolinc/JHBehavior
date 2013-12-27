using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

using Framework;

namespace JHSchool.Behavior
{
    /// <summary>
    /// 此類別代表學生的某一筆獎勵記錄。
    /// </summary>
    public class MeritRecord
    {

        /// <summary>
        /// Constructor 
        /// </summary>
        internal MeritRecord(string refStudentID, XmlElement element)
        {
            XmlHelper helper = new XmlHelper(element);

            RefStudentID = refStudentID;


            RefStudentID = refStudentID;
            ID = helper.GetString("@ID");                                               //獎懲編號
            SchoolYear = helper.GetString("SchoolYear");                                //學年度
            Semester = helper.GetString("Semester");                                    //學期
            GradeYear = helper.GetString("GradeYear");                                  //年級
            ClassName = helper.GetString("ClassName");                                  //班級名稱
            SeatNo = helper.GetString("SeatNo");
            Gender = helper.GetString("Gender");                                        //姓別
            OccurDate = helper.GetDateString("OccurDate");                              //獎勵日期
            Type = helper.GetString("Type");
            StudentNumber = helper.GetString("StudentNumber");                          //學生姓名?
            Reason = helper.GetString("Reason");                                        //事由
            RegisterDate = helper.GetString("RegisterDate");                            //登錄日期
            MeritA = helper.GetString("Detail/Discipline/Merit/@A");
            MeritB = helper.GetString("Detail/Discipline/Merit/@B");
            MeritC = helper.GetString("Detail/Discipline/Merit/@C");
            MeritFlag = helper.GetString("MeritFlag");                                  //0是懲戒,1是獎勵,2是留察

        }

        #region ========= Properties ========

        internal string RefStudentID { get; private set; }
        public string ID { get; private set; }
        public string Semester { get; private set; }
        public string OccurDate { get; private set; }

        public string Type { get; private set; }
        public string StudentNumber { get; private set; }
        public string SchoolYear { get; private set; }

        public string GradeYear { get; private set; }
        public string MeritFlag { get; private set; }
        public string Name { get; private set; }

        public string SeatNo { get; private set; }
        public string Gender { get; private set; }
        public string ClassName { get; private set; }
        public string Reason { get; private set; }
        public string RegisterDate { get; private set; } //登錄日期

        public string MeritA { get; private set; }    //大功數
        public string MeritB { get; private set; }    //小功數
        public string MeritC { get; private set; }    //獎勵數

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