using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JHSchool.Data;

namespace Behavior.ClubActivitiesPointList
{
    class GetCourseDetail
    {
        private List<JHCourseRecord> _CourseList = new List<JHCourseRecord>(); //課程
        private List<string> _CourseKeyList = new List<string>(); //課程ID
        private List<JHSCAttendRecord> SCAttendList = new List<JHSCAttendRecord>(); //課程修課學生
        /// <summary>
        /// 取得"課程ID,修課學生清單"
        /// </summary>
        public Dictionary<string, List<JHStudentRecord>> DicCourseStudent = new Dictionary<string, List<JHStudentRecord>>();
        private Dictionary<string, List<JHStudentRecord>> SortCourseStudent = new Dictionary<string, List<JHStudentRecord>>();
        /// <summary>
        /// 以課程ID取得課程詳細資訊
        /// </summary>
        public Dictionary<string, JHCourseRecord> DicCourseByKey = new Dictionary<string, JHCourseRecord>();

        //建構子
        public GetCourseDetail(List<JHCourseRecord> CourseList)
        {
            _CourseList = CourseList;

            foreach (JHCourseRecord each in _CourseList)
            {
                if (!_CourseKeyList.Contains(each.ID))
                {
                    _CourseKeyList.Add(each.ID);
                }
            }

            List<JHSCAttendRecord> SCAttendList = JHSCAttend.SelectByStudentIDAndCourseID(new List<string>(), _CourseKeyList);
            //課程排序用
            List<JHCourseRecord> SortCourse = new List<JHCourseRecord>();
            //轉回課程Key
            List<string> CourseKeyList = new List<string>();
            foreach (JHSCAttendRecord each in SCAttendList)
            {
                if (!SortCourseStudent.ContainsKey(each.RefCourseID))
                {
                    SortCourseStudent.Add(each.RefCourseID, new List<JHStudentRecord>());
                    DicCourseByKey.Add(each.RefCourseID, each.Course);
                    //課程排序用
                    SortCourse.Add(each.Course);
                }

                SortCourseStudent[each.RefCourseID].Add(each.Student);
            }

            //課程排序用
            SortCourse.Sort(new Comparison<JHCourseRecord>(CourseComparer));
            //轉回課程Key
            foreach (JHCourseRecord each in SortCourse)
            {
                CourseKeyList.Add(each.ID);
            }

            foreach (string each in CourseKeyList)
            {
                //學生排序用
                List<JHStudentRecord> list = SortCourseStudent[each];
                //學生排序
                list.Sort(new Comparison<JHStudentRecord>(StudentComparer));
                //組合回去為資料內容
                DicCourseStudent.Add(each, list);
            }
        }


        public static int CourseComparer(JHCourseRecord x, JHCourseRecord y)
        {
            string xx = x.Name;
            string yy = y.Name;

            return xx.CompareTo(yy);
        }


        public static int StudentComparer(JHStudentRecord x, JHStudentRecord y)
        {
            string Xcheck;
            if (x.Class != null)
            {
                Xcheck = x.Class.Name;
            }
            else
            {
                Xcheck = "00000";
            }
            string Ycheck;
            if (y.Class != null)
            {
                Ycheck = y.Class.Name;
            }
            else
            {
                Ycheck = "00000";
            }
            
            int xx = x.SeatNo.HasValue ? x.SeatNo.Value : 0;
            int yy = y.SeatNo.HasValue ? y.SeatNo.Value : 0;

            Xcheck += xx.ToString().PadLeft(5, '0');
            Ycheck += yy.ToString().PadLeft(5, '0');


            return Xcheck.CompareTo(Ycheck);
        }
    }
}
