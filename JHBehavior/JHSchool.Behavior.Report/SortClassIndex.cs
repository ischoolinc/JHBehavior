using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JHSchool.Behavior.Report
{
    public static class SortClassIndex
    {

        #region JHClassRecord / JHStudentRecord

        static public List<JHSchool.Data.JHClassRecord> JHSchoolData_JHClassRecord(List<JHSchool.Data.JHClassRecord> ClassList)
        {
            ClassList.Sort(SortJHSchoolData_JHClassRecord);
            return ClassList;
        }

        static private int SortJHSchoolData_JHClassRecord(JHSchool.Data.JHClassRecord class1, JHSchool.Data.JHClassRecord class2)
        {
            string ClassYear1 = class1.GradeYear.HasValue ? class1.GradeYear.Value.ToString().PadLeft(10, '0') : string.Empty.PadLeft(10, '9');
            string ClassYear2 = class2.GradeYear.HasValue ? class2.GradeYear.Value.ToString().PadLeft(10, '0') : string.Empty.PadLeft(10, '9');

            string DisplayOrder1 = "";
            if (string.IsNullOrEmpty(class1.DisplayOrder))
            {
                DisplayOrder1 = class1.DisplayOrder.PadLeft(10, '9');
            }
            else
            {
                DisplayOrder1 = class1.DisplayOrder.PadLeft(10, '0');
            }
            string DisplayOrder2 = "";
            if (string.IsNullOrEmpty(class2.DisplayOrder))
            {
                DisplayOrder2 = class2.DisplayOrder.PadLeft(10, '9');
            }
            else
            {
                DisplayOrder2 = class2.DisplayOrder.PadLeft(10, '0');
            }

            string ClassName1 = class1.Name.PadLeft(10, '0');
            string ClassName2 = class2.Name.PadLeft(10, '0');

            string Compareto1 = ClassYear1 + DisplayOrder1 + ClassName1;
            string Compareto2 = ClassYear2 + DisplayOrder2 + ClassName2;

            return Compareto1.CompareTo(Compareto2);
        }

        static public List<JHSchool.Data.JHStudentRecord> JHSchoolData_JHStudentRecord(List<JHSchool.Data.JHStudentRecord> StudentList)
        {
            //整理出學生&班級資料清單
            List<string> classIDList = new List<string>();
            foreach (JHSchool.Data.JHStudentRecord student in StudentList)
            {
                if (!string.IsNullOrEmpty(student.RefClassID))
                {
                    if (!classIDList.Contains(student.RefClassID))
                    {
                        classIDList.Add(student.RefClassID);
                    }
                }
            }
            //一次取得班級清單
            List<JHSchool.Data.JHClassRecord> classList = JHSchool.Data.JHClass.SelectByIDs(classIDList);
            //班級ID對照清單
            Dictionary<string,JHSchool.Data.JHClassRecord> classDic = new Dictionary<string,Data.JHClassRecord>();
            foreach (JHSchool.Data.JHClassRecord classRecord in classList)
            {
                if (!classDic.ContainsKey(classRecord.ID))
                {
                    classDic.Add(classRecord.ID, classRecord);
                }
            }

            List<StudentSortObj_JHSchoolData> list = new List<StudentSortObj_JHSchoolData>();
            foreach (JHSchool.Data.JHStudentRecord student in StudentList)
            {
                if (!string.IsNullOrEmpty(student.RefClassID))
                {
                    StudentSortObj_JHSchoolData obj = new StudentSortObj_JHSchoolData(classDic[student.RefClassID],student);
                    list.Add(obj);
                }
                else
                {
                    StudentSortObj_JHSchoolData obj = new StudentSortObj_JHSchoolData(null, student);
                    list.Add(obj);
                }
            }
            list.Sort(SortJHSchoolData_JHStudentRecord);

            return list.Select(x => x._StudentRecord).ToList();

        }

        static private int SortJHSchoolData_JHStudentRecord(StudentSortObj_JHSchoolData obj1, StudentSortObj_JHSchoolData obj2)
        {
            return obj1._SortString.CompareTo(obj2._SortString);
        }
        
        #endregion

        #region K12ClassRecord

        static public List<K12.Data.ClassRecord> K12Data_ClassRecord(List<K12.Data.ClassRecord> ClassList)
        {
            ClassList.Sort(SortK12Data_ClassRecord);
            return ClassList;
        }

        static private int SortK12Data_ClassRecord(K12.Data.ClassRecord class1, K12.Data.ClassRecord class2)
        {
            string ClassYear1 = class1.GradeYear.HasValue ? class1.GradeYear.Value.ToString().PadLeft(10, '0') : string.Empty.PadLeft(10, '9');
            string ClassYear2 = class2.GradeYear.HasValue ? class2.GradeYear.Value.ToString().PadLeft(10, '0') : string.Empty.PadLeft(10, '9');

            string DisplayOrder1 = "";
            if (string.IsNullOrEmpty(class1.DisplayOrder))
            {
                DisplayOrder1 = class1.DisplayOrder.PadLeft(10, '9');
            }
            else
            {
                DisplayOrder1 = class1.DisplayOrder.PadLeft(10, '0');
            }
            string DisplayOrder2 = "";
            if (string.IsNullOrEmpty(class2.DisplayOrder))
            {
                DisplayOrder2 = class2.DisplayOrder.PadLeft(10, '9');
            }
            else
            {
                DisplayOrder2 = class2.DisplayOrder.PadLeft(10, '0');
            }

            string ClassName1 = class1.Name.PadLeft(10, '0');
            string ClassName2 = class2.Name.PadLeft(10, '0');

            string Compareto1 = ClassYear1 + DisplayOrder1 + ClassName1;
            string Compareto2 = ClassYear2 + DisplayOrder2 + ClassName2;

            return Compareto1.CompareTo(Compareto2);
        }

        static public List<K12.Data.StudentRecord> K12Data_StudentRecord(List<K12.Data.StudentRecord> StudentList)
        {
            //整理出學生&班級資料清單
            List<string> classIDList = new List<string>();
            foreach (K12.Data.StudentRecord student in StudentList)
            {
                if (!string.IsNullOrEmpty(student.RefClassID))
                {
                    if (!classIDList.Contains(student.RefClassID))
                    {
                        classIDList.Add(student.RefClassID);
                    }
                }
            }
            //一次取得班級清單
            List<K12.Data.ClassRecord> classList = K12.Data.Class.SelectByIDs(classIDList);
            //班級ID對照清單
            Dictionary<string, K12.Data.ClassRecord> classDic = new Dictionary<string, K12.Data.ClassRecord>();
            foreach (K12.Data.ClassRecord classRecord in classList)
            {
                if (!classDic.ContainsKey(classRecord.ID))
                {
                    classDic.Add(classRecord.ID, classRecord);
                }
            }

            List<StudentSortObj_K12Data> list = new List<StudentSortObj_K12Data>();
            foreach (K12.Data.StudentRecord student in StudentList)
            {
                if (!string.IsNullOrEmpty(student.RefClassID))
                {
                    StudentSortObj_K12Data obj = new StudentSortObj_K12Data(classDic[student.RefClassID], student);
                    list.Add(obj);
                }
                else
                {
                    StudentSortObj_K12Data obj = new StudentSortObj_K12Data(null, student);
                    list.Add(obj);
                }
            }
            list.Sort(SortK12Data_StudentRecord);

            return list.Select(x => x._StudentRecord).ToList();

        }

        static private int SortK12Data_StudentRecord(StudentSortObj_K12Data obj1, StudentSortObj_K12Data obj2)
        {
            return obj1._SortString.CompareTo(obj2._SortString);
        }
        #endregion

        static public List<JHSchool.ClassRecord> JHSchool_ClassRecord(List<JHSchool.ClassRecord> ClassList)
        {
            ClassList.Sort(SortJHScool_ClassRecord);
            return ClassList;
        }

        static private int SortJHScool_ClassRecord(JHSchool.ClassRecord class1, JHSchool.ClassRecord class2)
        {
            string ClassYear1 = !string.IsNullOrEmpty(class1.GradeYear) ? class1.GradeYear.PadLeft(10, '0') : class1.GradeYear.PadLeft(10, '9');
            string ClassYear2 = !string.IsNullOrEmpty(class2.GradeYear) ? class2.GradeYear.PadLeft(10, '0') : class2.GradeYear.PadLeft(10, '9');

            string DisplayOrder1 = "";
            if (string.IsNullOrEmpty(class1.DisplayOrder))
            {
                DisplayOrder1 = class1.DisplayOrder.PadLeft(10, '9');
            }
            else
            {
                DisplayOrder1 = class1.DisplayOrder.PadLeft(10, '0');
            }
            string DisplayOrder2 = "";
            if (string.IsNullOrEmpty(class2.DisplayOrder))
            {
                DisplayOrder2 = class2.DisplayOrder.PadLeft(10, '9');
            }
            else
            {
                DisplayOrder2 = class2.DisplayOrder.PadLeft(10, '0');
            }

            string ClassName1 = class1.Name.PadLeft(10, '0');
            string ClassName2 = class2.Name.PadLeft(10, '0');

            string Compareto1 = ClassYear1 + DisplayOrder1 + ClassName1;
            string Compareto2 = ClassYear2 + DisplayOrder2 + ClassName2;

            return Compareto1.CompareTo(Compareto2);
        }

        #region JHSchool(註解)

        //依"班級序號/班級名稱/學生座號/學生姓名"排序
        static public List<JHSchool.StudentRecord> JHSchool_StudentRecord(List<JHSchool.StudentRecord> StudentList)
        {
            //取得"班級"Record未完成
            List<StudentSortObj_JHSchool> list = new List<StudentSortObj_JHSchool>();
            foreach (JHSchool.StudentRecord student in StudentList)
            {
                StudentSortObj_JHSchool obj = new StudentSortObj_JHSchool(student);
                list.Add(obj);
            }
            list.Sort(SortJHScool_StudentRecord);

            return list.Select(x => x._StudentRecord).ToList();

        }

        static private int SortJHScool_StudentRecord(StudentSortObj_JHSchool obj1, StudentSortObj_JHSchool obj2)
        {
            return obj1._SortString.CompareTo(obj2._SortString);
        }
        #endregion
    }
}
