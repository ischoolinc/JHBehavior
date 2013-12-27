using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JHSchool.Data;
using JHSchool.Logic;

namespace JHSchool.Behavior
{
    class SortClass
    {
        public SortClass()
        {

        }

        /// <summary>
        /// 傳入懲戒資料,依學生班級座號排序
        /// </summary>
        public int SortDemeritRecord(JHDemeritRecord x, JHDemeritRecord y)
        {
            JHStudentRecord student1 = x.Student;
            JHStudentRecord student2 = y.Student;

            return SortStudent(student1, student2);
        }

        /// <summary>
        /// 傳入獎勵資料,依學生班級座號排序
        /// </summary>
        public int SortMeritRecord(JHMeritRecord x, JHMeritRecord y)
        {
            JHStudentRecord student1 = x.Student;
            JHStudentRecord student2 = y.Student;

            return SortStudent(student1, student2);
        }

        /// <summary>
        /// 傳入缺曠,依學生班級座號排序
        /// </summary>
        public int SortAttendanceRecord(JHAttendanceRecord x, JHAttendanceRecord y)
        {
            JHStudentRecord student1 = x.Student;
            JHStudentRecord student2 = y.Student;

            return SortStudent(student1, student2);
        }

        /// <summary>
        /// 傳入自動統計,依學生班級座號排序(有速度問題?)
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public int SortAutoSummaryRecord(AutoSummaryRecord x, AutoSummaryRecord y)
        {
            JHStudentRecord student1 = x.Student;
            JHStudentRecord student2 = y.Student;

            return SortStudent(student1, student2);
        }

        /// <summary>
        /// 傳入學生,依學生班級座號排序
        /// </summary>
        public int SortStudent(JHStudentRecord x, JHStudentRecord y)
        {
            JHStudentRecord student1 = x;
            JHStudentRecord student2 = y;

            string ClassName1 = student1.Class != null ? student1.Class.Name : "";
            ClassName1 = ClassName1.PadLeft(5, '0');
            string ClassName2 = student2.Class != null ? student2.Class.Name : "";
            ClassName2 = ClassName2.PadLeft(5, '0');

            string Sean1 = student1.SeatNo.HasValue ? student1.SeatNo.Value.ToString() : "";
            Sean1 = Sean1.PadLeft(3, '0');
            string Sean2 = student2.SeatNo.HasValue ? student2.SeatNo.Value.ToString() : "";
            Sean2 = Sean2.PadLeft(3, '0');

            ClassName1 += Sean1;
            ClassName2 += Sean2;

            return ClassName1.CompareTo(ClassName2);
        }
        
    }
}
