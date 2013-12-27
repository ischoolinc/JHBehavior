using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using JHSchool.Editor;

namespace JHSchool.Behavior.Editor
{
    public class MeritRecordEditor
    {
        internal MeritRecordEditor(MeritRecord meritRecord)
        {
            MeritRecord = meritRecord;
            
            Remove = false;

            RefStudentID = meritRecord.RefStudentID;
            ID = meritRecord.ID;
            GradeYear = meritRecord.GradeYear;
            ClassName = meritRecord.ClassName;
            Gender = meritRecord.Gender;
            MeritA = meritRecord.MeritA;
            MeritB = meritRecord.MeritB;
            MeritC = meritRecord.MeritC;
            MeritFlag = meritRecord.MeritFlag;
            Name = meritRecord.Name;
            OccurDate = meritRecord.OccurDate;
            Reason = meritRecord.Reason;
            RegisterDate = meritRecord.RegisterDate;
            SchoolYear = meritRecord.SchoolYear;
            SeatNo = meritRecord.SeatNo;
            Semester = meritRecord.Semester;
            StudentNumber = meritRecord.StudentNumber;
            Type = meritRecord.Type;           
        }

        public MeritRecordEditor(StudentRecord studentRecord)
        {
            RefStudentID = studentRecord.ID;
            GradeYear = studentRecord.Class.GradeYear;
        }

        public EditorStatus EditorStatus
        {
            get
            {
                if (MeritRecord == null)
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
            }
        }

        public void Save()
        {
            if (this.EditorStatus != EditorStatus.NoChanged)
                Feature.EditMerit.SaveMeritRecordEditor(this);
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
        /// <summary>
        /// 0是懲戒,1是獎勵,2是留察
        /// </summary>
        public string MeritFlag {get;set;}

        /// <summary>
        /// 事由
        /// </summary>
        public string Reason { get; set; }
        /// <summary>
        /// 登錄日期
        /// </summary>
        public string RegisterDate { get; set; }
        /// <summary>
        /// 大功數
        /// </summary>
        public string MeritA { get; set; }
        /// <summary>
        /// 小功數
        /// </summary>
        public string MeritB { get; set; }
        /// <summary>
        /// 獎勵數
        /// </summary>
        public string MeritC { get; set; }

        public StudentRecord Student
        {
            get { return JHSchool.Student.Instance[RefStudentID]; }
        }

        #endregion

        internal MeritRecord MeritRecord{ get; private set; }
    }

    public static class MeritRecordEditor_ExtendMethods
    {
        public static MeritRecordEditor GetEditor(this MeritRecord meritRecord)
        {
            return new MeritRecordEditor(meritRecord);
        }

        public static void SaveAll(this IEnumerable<MeritRecordEditor> editors)
        {
            Feature.EditMerit.SaveMeritRecordEditors(editors);
        }

        public static MeritRecordEditor AddDemeritRecord(this StudentRecord studentRec)
        {
            return new MeritRecordEditor(studentRec);
        }
    }
}