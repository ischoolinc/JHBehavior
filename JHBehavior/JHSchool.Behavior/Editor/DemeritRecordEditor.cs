using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JHSchool.Editor;
using Framework;

namespace JHSchool.Behavior.Editor
{
    public class DemeritRecordEditor
    {
        private string meritFlag = "";
        private bool isInObservation = false;

        /// <summary>
        /// Constructor，為修改模式使用
        /// </summary>
        /// <param name="demeritRecord"></param>
        internal DemeritRecordEditor(DemeritRecord demeritRecord)
        {
            DemeritRecord = demeritRecord;    //避免 Cicular Reference 
            Remove = false;
            RefStudentID = demeritRecord.RefStudentID;
            ID = demeritRecord.ID;

            SchoolYear = demeritRecord.SchoolYear;                                //學年度
            Semester = demeritRecord.Semester;                              //學期
            GradeYear = demeritRecord.GradeYear;                                  //年級
            ClassName = demeritRecord.ClassName;                                  //班級名稱
            SeatNo = demeritRecord.SeatNo;                                        //
            Gender = demeritRecord.Gender;                                        //姓別
            OccurDate = demeritRecord.OccurDate;                              //懲戒日期
            Type = demeritRecord.Type;                                            // ??? 
            StudentNumber = demeritRecord.StudentNumber;                          //學生姓名
            Reason = demeritRecord.Reason;                                        //事由
            RegisterDate = demeritRecord.RegisterDate;                          //登錄日期

            DemeritA = demeritRecord.DemeritA;                //大過
            DemeritB = demeritRecord.DemeritB;                //小過
            DemeritC = demeritRecord.DemeritC;                //警告  

            ClearDate = demeritRecord.ClearDate;       //銷過日期
            ClearReason = demeritRecord.ClearReason;   //銷過事由
            Cleared = demeritRecord.Cleared;           //銷過

            IsCleared = !string.IsNullOrEmpty(demeritRecord.Cleared);                      //是否銷過
            
            meritFlag = demeritRecord.MeritFlag;                                  //0是懲戒,1是獎勵,2是留察
            IsInObservation = (MeritFlag == "2");           //是否留校察看


        }

        /// <summary>
        /// Constructor ，為新增模式使用。
        /// </summary>
        /// <param name="studentRecord"></param>
        public DemeritRecordEditor(StudentRecord studentRecord)
        {
            Attributes = new AutoDictionary();
            DemeritRecord = null;
            Remove = false;
            RefStudentID = studentRecord.ID;
            GradeYear = studentRecord.Class.GradeYear;
        }

        /// <summary>
        /// 取得修改狀態
        /// </summary>
        public EditorStatus EditorStatus
        {
            get
            {
                if (DemeritRecord == null)
                {
                    if (!Remove)
                        return EditorStatus.Insert;
                    else
                        return EditorStatus.NoChanged;
                }
                else
                {
                    if (Remove)
                        return  EditorStatus.Delete;
                    else
                        return EditorStatus.Update;
                    
                }

                //return EditorStatus.NoChanged;
            }
        }

        public virtual void Save()
        {
            if (this.EditorStatus != EditorStatus.NoChanged)
                Feature.EditDemerit.SaveDemeritRecordEditor(this);                
        }

        #region Fields

        /// <summary>
        /// 刪除請設成 False。
        /// </summary>
        public bool Remove { get; set; }
        /// <summary>
        /// 不可以修改。
        /// </summary>
        internal string RefStudentID { get; private set; }
        /// <summary>
        /// 不可以修改。
        /// </summary>
        public string ID { get; private set; }

        
        /// <summary>
        /// 年級
        /// </summary>
        public string GradeYear { get; private set; }
        
        /// <summary>
        /// 日期
        /// </summary>
        public string OccurDate { get; set; }

        public string Type { get; private set; }
        public string StudentNumber { get; private set; }
        public string Name { get; private set; }
        public string ClassName { get; private set; }
        public string SeatNo { get; private set; }
        public string Gender { get; private set; }        

        /// <summary>
        /// 學年度
        /// </summary>
        public string SchoolYear { get; set; }
        /// <summary>
        /// 學期
        /// </summary>
        public string Semester { get; set; }


        //public string OriginDepartment { get; private set; }
        /// <summary>
        /// 0是懲戒,1是獎勵,2是留察
        /// </summary>
        public string MeritFlag {
            get
            {
                return this.meritFlag;
            }
            set
            {
                this.meritFlag = value;
                this.isInObservation = (this.MeritFlag =="2");
            }
        }
        /// <summary>
        /// 是否留校察看
        /// </summary>
        public bool IsInObservation {
            get { return this.isInObservation; }
            set { 
                this.isInObservation = value;
                if (this.IsInObservation)
                {
                    this.DemeritA = "0";
                    this.DemeritB = "0";
                    this.DemeritC = "0";
                    this.meritFlag = "2";
                }
            }
        }

        /// <summary>
        /// 事由
        /// </summary>
        public string Reason { get; set; }
        /// <summary>
        /// 登錄日期
        /// </summary>
        public string RegisterDate { get; set; }

        /// <summary>
        /// 大過數
        /// </summary>
        public string DemeritA { get; set; }
        /// <summary>
        /// 小過數
        /// </summary>
        public string DemeritB { get; set; }
        /// <summary>
        /// 警告數
        /// </summary>
        public string DemeritC { get; set; }
        /// <summary>
        /// 是否銷過。加這個屬性是在修改時，判斷是否有銷過資料。
        /// </summary>
        public bool IsCleared { get; set; }
        /// <summary>
        /// 銷過日期
        /// </summary>
        public string ClearDate { get; set; }
        /// <summary>
        /// 銷過事由
        /// </summary>
        public string ClearReason { get; set; }
        /// <summary>
        /// 是否銷過
        /// </summary>
        public string Cleared { get; set; }



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

        public AutoDictionary Attributes { get; private set; }

        public StudentRecord Student
        {
            get { return JHSchool.Student.Instance[RefStudentID]; }
        }

        #endregion

        internal DemeritRecord DemeritRecord { get; private set; } 
    }

    public static class DemeritRecordEditor_ExtendFunctions
    {
        public static DemeritRecordEditor GetEditor(this DemeritRecord demeritRecord)
        {
            return new DemeritRecordEditor(demeritRecord);
        }

        public static void SaveAll(this IEnumerable<DemeritRecordEditor> editors)
        {
            Feature.EditDemerit.SaveDemeritRecordEditors(editors);
        }

        public static DemeritRecordEditor AddDemeritRecord(this StudentRecord studentRec)
        {
            return new DemeritRecordEditor(studentRec);
        }
    }
}
