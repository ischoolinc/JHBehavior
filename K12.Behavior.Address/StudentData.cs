using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Presentation.Controls;
using K12.Data;

namespace K12.Behavior.Address
{
    public class StudentData
    {
        //班級清單
        private List<ClassRecord> _ClassList = new List<ClassRecord>();
        //學生清單
        private List<StudentRecord> _StudRecordList = new List<StudentRecord>();
        //學生ID
        private List<string> _StudentIdList = new List<string>();

        private Dictionary<string, AddressRecord> _Address = new Dictionary<string, AddressRecord>();

        private Dictionary<string, PhoneRecord> _Phone = new Dictionary<string, PhoneRecord>();

        public StudentData()
        {
            //List<StudentRecord> Test = JHStudent.SelectAll();
            Reset();
        }

        public void Reset()
        {
            //int t = Environment.TickCount;
            GetStudent(); //取得學生資料
            //MsgBox.Show("" + (Environment.TickCount - t));
            //t = Environment.TickCount;
            GetAddress(); //取得地址資料
            //MsgBox.Show("" + (Environment.TickCount - t));
            //t = Environment.TickCount;
            GetPhone(); //取得電話資料
            //MsgBox.Show("" + (Environment.TickCount - t));
            //t = Environment.TickCount;
        }

        #region 內部取得資料用
        private void GetStudent()
        {
            #region 取得學生資料
            _ClassList.Clear();
            _StudRecordList.Clear();
            _StudentIdList.Clear();
            _Address.Clear();
            _Phone.Clear();

            //_ClassList = JHClass.SelectByIDs(Class.Instance.SelectedKeys);
            //List<string> ChengClassIDs = new List<string>(); //轉換成ID
            //foreach (ClassRecord each in _ClassList)
            //{
            //    ChengClassIDs.Add(each.ID);
            //}

            List<StudentRecord> AllStudentReocrd = Student.SelectByClasses(Class.SelectByIDs(K12.Presentation.NLDPanels.Class.SelectedSource));
            AllStudentReocrd = SortClassIndex.K12Data_StudentRecord(AllStudentReocrd); //排序

            foreach (StudentRecord stud in AllStudentReocrd)
            {
                if (CheckStatus(stud))
                {
                    if (!_Address.ContainsKey(stud.ID))
                    {
                        _Address.Add(stud.ID, null); //地址
                        _Phone.Add(stud.ID, null); //電話
                        _StudentIdList.Add(stud.ID); //ID
                        _StudRecordList.Add(stud); //Record
                    }
                }
            }
            

            //foreach (ClassRecord each in _ClassList)
            //{ 

            //    List<StudentRecord> studList = each.Students;

            //    studList.Sort(new Comparison<StudentRecord>(SortSeatNo)); //排序

            //    foreach (StudentRecord stud in studList)
            //    {
            //        if (stud.Status == K12.Data.StudentRecord.StudentStatus.一般)
            //        {
            //            if (!_Address.ContainsKey(stud.ID))
            //            {
            //                _Address.Add(stud.ID, null); //地址
            //                _Phone.Add(stud.ID, null); //電話
            //                _StudentIdList.Add(stud.ID); //ID
            //                _StudRecordList.Add(stud); //Record
            //            }
            //        }
            //    }
            //}

            #endregion
        }

        private bool CheckStatus(StudentRecord stud)
        {
            if (stud.Status == StudentRecord.StudentStatus.一般 || stud.Status == StudentRecord.StudentStatus.延修)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
                 
        private void GetAddress()
        {
            #region 取得地址資料
            List<AddressRecord> list = K12.Data.Address.SelectByStudentIDs(_StudentIdList);
            foreach (AddressRecord each in list)
            {
                if (_Address.ContainsKey(each.RefStudentID))
                {
                    _Address[each.RefStudentID] = each;
                }
            } 
            #endregion
        }

        private void GetPhone()
        {
            #region 取得電話資料
            List<PhoneRecord> list = K12.Data.Phone.SelectByStudentIDs(_StudentIdList);

            foreach (PhoneRecord each in list)
            {
                if (_Phone.ContainsKey(each.RefStudentID))
                {
                    _Phone[each.RefStudentID] = each;
                }
            } 
            #endregion
        } 
        #endregion

        #region 對外資料
        /// <summary>
        /// 取得學生資料
        /// </summary>
        public List<StudentRecord> StudRecordList
        {
            get { return _StudRecordList; }
            set { _StudRecordList = StudRecordList; }
        }

        /// <summary>
        /// 取得地址清單
        /// </summary>
        public Dictionary<string, AddressRecord> Address
        {
            get { return _Address; }
            set { _Address = Address; }
        }

        /// <summary>
        /// 取得電話清單
        /// </summary>
        public Dictionary<string, PhoneRecord> Phone
        {
            get { return _Phone; }
            set { _Phone = Phone; }
        }

        #endregion

        #region 排序
        private int SortSeatNo(StudentRecord x, StudentRecord y)
        {
            string xx1 = x.Class.Name;
            string xx2 = x.SeatNo.HasValue ? x.SeatNo.Value.ToString().PadLeft(3, '0') : "000";
            string xx3 = xx1 + xx2;

            string yy1 = y.Class.Name;
            string yy2 = y.SeatNo.HasValue ? y.SeatNo.Value.ToString().PadLeft(3, '0') : "000";
            string yy3 = yy1 + yy2;

            return xx3.CompareTo(yy3);

            //int xx = x.SeatNo.HasValue ? x.SeatNo.Value : 0;
            //int yy = y.SeatNo.HasValue ? y.SeatNo.Value : 0;
            //return xx.CompareTo(yy);
        }
        //班級名稱用
        private int SortClassName(ClassRecord x, ClassRecord y)
        {
            string xx = x.Name;
            string yy = y.Name;
            return xx.CompareTo(yy);
        }
        #endregion
    }
}
