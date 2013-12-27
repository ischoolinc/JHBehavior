using System;
using System.Collections.Generic;
using System.Text;
using Framework;
using System.ComponentModel;
using FISCA.DSAUtil;
using System.Xml;
using FISCA.Authentication;

namespace JHSchool.Behavior
{
    /// <summary>
    /// 提供根據條件取得缺曠記錄，並暫存資料的類別。
    /// 此類別不作 Cache ，用後即丟。
    /// </summary>
    public class AttendanceFinder : DataCollection<AttendanceRecord>
    {
        private const string GET_ATTENDANCE = "SmartSchool.Student.Attendance.GetAttendance";
        private BackgroundWorker _backWorker;
        private int _schoolyear;
        private int _semester;
        private string _startDate = "";
        private string _endDate = "";
        private List<StudentRecord> _students;

        //Declare an event
        public event EventHandler OnAttendanceRecordReload;

        //Constructor
        public AttendanceFinder(List<StudentRecord> students)
        {
            this._students = students;
        }

        /// <summary>
        /// 根據學年度學期取得缺曠資料。
        /// 此方法以非同步方式呼叫服務。完成後會觸發 OnAttendanceRecordReload 事件。
        /// </summary>
        /// <param name="schoolyear">學年度</param>
        /// <param name="semester">學期</param>
        public void GetAttendanceRecordsBySemesterAsync(int schoolyear, int semester)
        {
            this._schoolyear = schoolyear;
            this._semester = semester;

            this._backWorker = new BackgroundWorker();
            this._backWorker.DoWork += new DoWorkEventHandler(getData_bySemester);
            this._backWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_backWorker_RunWorkerCompleted);
            this._backWorker.RunWorkerAsync();
        }
        /// <summary>
        /// 根據學年度學期取得缺曠資料。
        /// 此方法以同步方式呼叫服務，資料呼叫期間，畫面會停住。
        /// </summary>
        /// <param name="schoolyear">學年度</param>
        /// <param name="semester">學期</param>
        public Dictionary<string, List<AttendanceRecord>> GetAttendanceRecordsBySemester(int schoolyear, int semester)
        {
            this.ResetCollection();

            if (this._students.Count == 0) return null ;

            StringBuilder req = new StringBuilder("<Request><Field><All/></Field><Condition>");
            foreach (StudentRecord sr in this._students)
            {
                req.Append("<RefStudentID>" + sr.ID + "</RefStudentID>");
            }
            req.Append("<SchoolYear>" + this._schoolyear + "</SchoolYear><Semester>" + this._semester + "</Semester></Condition>");
            req.Append("<Order><SchoolYear/><Semester/><OccurDate/></Order></Request>");

            this.GetData(req.ToString());

            return this.records;
        }

        /// <summary>
        /// 根據日期區間取得缺曠資料。
        /// 此方法以非同步方式呼叫服務。完成後會觸發 OnAttendanceRecordReload 事件。
        /// </summary>
        /// <param name="schoolyear">學年度</param>
        /// <param name="semester">學期</param>
        public void GetAttendanceRecordsByDateAsync(string startDate, string endDate)
        {
            this._startDate = startDate;
            this._endDate = endDate;

            this._backWorker = new BackgroundWorker();
            this._backWorker.DoWork += new DoWorkEventHandler(getData_byDate);
            this._backWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_backWorker_RunWorkerCompleted);
            this._backWorker.RunWorkerAsync();
        }
                /// <summary>
        /// 根據日期區間取得缺曠資料。
        /// 此方法以同步方式呼叫服務，資料呼叫期間，畫面會停住。
        /// </summary>
        /// <param name="schoolyear">學年度</param>
        /// <param name="semester">學期</param>
        public Dictionary<string, List<AttendanceRecord>> GetAttendanceRecordsByDate(string startDate, string endDate)
        {
            this.ResetCollection();

            if (this._students.Count == 0) return null;

            StringBuilder req = new StringBuilder("<Request><Field><All/></Field><Condition>");
            foreach (StudentRecord sr in this._students)
            {
                req.Append("<RefStudentID>" + sr.ID + "</RefStudentID>");
            }
            req.Append("<StartDate>" + this._startDate + "</StartDate><EndDate>" + this._endDate + "</EndDate></Condition>");
            req.Append("<Order><SchoolYear/><Semester/><OccurDate/></Order></Request>");

            this.GetData(req.ToString());

            return this.records;
        }

        /// <summary>
        /// 非同步呼叫完成後，會觸發 OnAttendanceRecordReload 事件。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void _backWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //this._callback(this, EventArgs.Empty);
            if (this.OnAttendanceRecordReload != null)
                this.OnAttendanceRecordReload(this, EventArgs.Empty);
        }

        void getData_bySemester(object sender, DoWorkEventArgs e)
        {
            this.GetAttendanceRecordsBySemester(this._schoolyear, this._semester);
        }

        void getData_byDate(object sender, DoWorkEventArgs e)
        {
            this.GetAttendanceRecordsByDate(this._startDate, this._endDate);
        }

        /// <summary>
        /// 重設記錄資料的集合內容
        /// </summary>
        private void ResetCollection()
        {
            this.Clear();

            foreach (StudentRecord sr in this._students)
            {
                this.Add(sr.ID, new List<AttendanceRecord>());
            }
        }

        /// <summary>
        /// 取得缺曠記錄，並回歸到所屬學生底下，並將關聯記錄於內部的集合物件中。
        /// </summary>
        /// <param name="req"></param>
        private void GetData(string req)
        {
            List<AttendanceRecord> result = new List<AttendanceRecord>();

            foreach (XmlElement item in DSAServices.CallService(GET_ATTENDANCE, new DSRequest(req.ToString())).GetContent().GetElements("Attendance"))
            {
                AttendanceRecord ar = new AttendanceRecord(item);
                this[ar.RefStudentID].Add(ar);
            }

        }
    }

    public static class AttendanceFinder_ExtensionMethods
    {
        /// <summary>
        /// 為學生集合物件加上一個延伸方法。
        /// 
        /// </summary>
        /// <example>
        /// 
        /// 
        ///  List<StudentRecord> studs = new List<StudentRecord>();
        ///  studs.Add(student);
        ///  AttendanceFinder attFinder = studs.GetAttendanceFinder();
        ///  List<AttendanceRecord> records = attFinder.GetAttendanceRecordsBySemester(97,1)[student.ID];
        ///  
        /// </example>
        /// <param name="students"></param>
        /// <returns></returns>
        public static AttendanceFinder GetAttendanceFinder(this List<StudentRecord> students)
        {
            return new AttendanceFinder(students);
        }

        /// <summary>
        /// 為學生物件加上一個延伸方法。
        /// </summary>
        /// <example>        /// 
        /// 
        ///  StudentRecord stud = Student.Instance.Items[this.PrimaryKey];
        ///  AttendanceFinder attFinder = stud.GetAttendanceFinder();
        ///  List<AttendanceRecord> records = attFinder.GetAttendanceRecordsBySemester(97,1)[stud.ID];
        ///  
        /// </example>
        /// <param name="students"></param>
        /// <returns></returns>
        public static AttendanceFinder GetAttendanceFinder(this StudentRecord student)
        {
            List<StudentRecord> studs = new List<StudentRecord>();
            studs.Add(student);
            return new AttendanceFinder(studs);
        }
    }
}
