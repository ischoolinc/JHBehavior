using System;
using System.Collections.Generic;
using JHSchool.Data;

namespace JHSchool.Behavior.Report.班級學生獎懲統計
{
    class InfoClass
    {
        ObjConfig _config;

        private List<string> studentList = new List<string>();
        private Dictionary<string, List<string>> ClassList = new Dictionary<string, List<string>>();

        /// <summary>
        /// 整理後的班級物件(班級ID,自定物件)
        /// </summary>
        public Dictionary<string, MeritDemeritInfo> ClassMeritDemeritList = new Dictionary<string, MeritDemeritInfo>();

        /// <summary>
        /// 班級清單
        /// </summary>
        public List<JHClassRecord> TempClass;

        public InfoClass(ObjConfig config)
        {
            _config = config;

            //建立studentListstudentList & ClassList
            SelectClassByStudentRecord();

            //取得選擇班級的學生缺曠記錄
            List<JHMeritRecord> MeritList = new List<JHMeritRecord>();
            if (!_config.InsertOrSetup)
                MeritList = JHMerit.SelectByOccurDate(studentList, _config.StartDate, _config.EndDate);
            else
                MeritList = JHMerit.SelectByRegisterDate(studentList, _config.StartDate, _config.EndDate);

            //整理以學生ID為單位清單
            Dictionary<string, List<JHMeritRecord>> MeritDic = new Dictionary<string, List<JHMeritRecord>>();

            foreach (JHMeritRecord merit in MeritList)
            {
                if (!MeritDic.ContainsKey(merit.RefStudentID))
                    MeritDic.Add(merit.RefStudentID, new List<JHMeritRecord>());
                MeritDic[merit.RefStudentID].Add(merit);
            }

            SetMeritInClass(MeritDic);


            List<JHDemeritRecord> DemrtieList = new List<JHDemeritRecord>();
            if (!_config.InsertOrSetup)
                DemrtieList = JHDemerit.SelectByOccurDate(studentList, _config.StartDate, _config.EndDate);
            else
                DemrtieList = JHDemerit.SelectByRegisterDate(studentList, _config.StartDate, _config.EndDate);

            //整理以學生ID為單位清單
            Dictionary<string, List<JHDemeritRecord>> DemeritDic = new Dictionary<string, List<JHDemeritRecord>>();

            foreach (JHDemeritRecord demerit in DemrtieList)
            {
                if (demerit.Cleared == "是") //如果是銷過記錄
                {
                    if (!_config.Cleared) //True為設定檔的"包含銷過"
                    {
                        continue;
                    }
                }

                if (!DemeritDic.ContainsKey(demerit.RefStudentID))
                    DemeritDic.Add(demerit.RefStudentID, new List<JHDemeritRecord>());
                DemeritDic[demerit.RefStudentID].Add(demerit);
            }

            SetDemeritInClass(DemeritDic);
        }

        /// <summary>
        /// 傳入獎,以建立資料
        /// </summary>
        /// <param name="merit"></param>
        public void SetMeritInClass(Dictionary<string, List<JHMeritRecord>> MeritDic)
        {
            foreach (string each in MeritDic.Keys) //取得學生id
            {
                //比對班級清單,取得班級物件
                MeritDemeritInfo obj = new MeritDemeritInfo();
                foreach (string classID in ClassList.Keys) //學生在哪一班
                {
                    if (ClassList[classID].Contains(each)) //在ClassList[each]班
                    {
                        obj = ClassMeritDemeritList[classID];
                        obj.MeritStudentCount++;
                    }
                }

                //取得某學生之獎勵資料
                foreach (JHMeritRecord merit in MeritDic[each]) //取得內容
                {
                    if (_config.SelectItems.Contains("大功"))
                    {
                        int meritA = merit.MeritA.HasValue ? merit.MeritA.Value : 0;
                        if (meritA != 0)
                        {
                            obj.MeritA += meritA;
                            obj.MeritAStudentCount++;
                        }
                    }

                    if (_config.SelectItems.Contains("小功"))
                    {
                        int meritB = merit.MeritB.HasValue ? merit.MeritB.Value : 0;
                        if (meritB != 0)
                        {
                            obj.MeritB += meritB;
                            obj.MeritBStudentCount++;
                        }
                    }

                    if (_config.SelectItems.Contains("嘉獎"))
                    {
                        int meritC = merit.MeritC.HasValue ? merit.MeritC.Value : 0;
                        if (meritC != 0)
                        {
                            obj.MeritC += meritC;
                            obj.MeritCStudentCount++;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 傳入懲,以建立資料
        /// </summary>
        /// <param name="merit"></param>
        public void SetDemeritInClass(Dictionary<string, List<JHDemeritRecord>> DemeritDic)
        {
            foreach (string each in DemeritDic.Keys) //取得學生id
            {
                MeritDemeritInfo obj = new MeritDemeritInfo();

                foreach (string classID in ClassList.Keys) //學生在哪一班
                {
                    if (ClassList[classID].Contains(each)) //在ClassList[each]班
                    {
                        obj = ClassMeritDemeritList[classID];
                        obj.DemeritStudentCount++;
                    }
                }

                foreach (JHDemeritRecord demerit in DemeritDic[each]) //取得內容
                {
                    if (_config.SelectItems.Contains("大過"))
                    {
                        int demeritA = demerit.DemeritA.HasValue ? demerit.DemeritA.Value : 0;
                        if (demeritA != 0)
                        {
                            obj.DemeritA += demeritA; //大過支數
                            obj.DemeritAStudentCount++; //大過人數
                        }
                    }

                    if (_config.SelectItems.Contains("小過"))
                    {
                        int demeritB = demerit.DemeritB.HasValue ? demerit.DemeritB.Value : 0;
                        if (demeritB != 0)
                        {
                            obj.DemeritB += demeritB;
                            obj.DemeritBStudentCount++;
                        }
                    }

                    if (_config.SelectItems.Contains("警告"))
                    {
                        int demeritC = demerit.DemeritC.HasValue ? demerit.DemeritC.Value : 0;
                        if (demeritC != 0)
                        {
                            obj.DemeritC += demeritC;
                            obj.DemeritCStudentCount++;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 依所有學生資訊建立班級清單
        /// </summary>
        public void SelectClassByStudentRecord()
        {
            //全部的班級
            Dictionary<string, List<string>> dic = new Dictionary<string, List<string>>();
            foreach (JHStudentRecord student in JHStudent.SelectAll())
            {
                if (student.Class == null) //沒有班級
                    continue;

                if (student.Status == K12.Data.StudentRecord.StudentStatus.一般 || student.Status == K12.Data.StudentRecord.StudentStatus.輟學) //只統計一般生 , 20120510-增加輟學
                {
                    if (!dic.ContainsKey(student.Class.ID)) //依班級編號建立清單
                    {
                        dic.Add(student.Class.ID, new List<string>());
                    }

                    dic[student.Class.ID].Add(student.ID);
                }
            }

            //選擇的班級
            studentList.Clear();
            TempClass = JHClass.SelectByIDs(K12.Presentation.NLDPanels.Class.SelectedSource);
            TempClass = SortClassIndex.JHSchoolData_JHClassRecord(TempClass);
            foreach (JHClassRecord each in TempClass)
            {
                if (!dic.ContainsKey(each.ID))
                    continue;

                ClassMeritDemeritList.Add(each.ID, new MeritDemeritInfo());
                ClassMeritDemeritList[each.ID].ClassName = each.Name;
                if (each.Teacher != null)
                {
                    ClassMeritDemeritList[each.ID].TeacherName = each.Teacher.Name; //可能造成速度問題
                }
                ClassMeritDemeritList[each.ID].StudentCount = dic[each.ID].Count;


                //建立清單
                if (dic.ContainsKey(each.ID))
                {
                    studentList.AddRange(dic[each.ID]); //學生清單
                    ClassList.Add(each.ID, dic[each.ID]); //依班級ID的學生清單
                }
            }
        }

        private int SortClass(JHClassRecord x, JHClassRecord y)
        {
            string xx = "";
            if (x.DisplayOrder == "")
                xx = x.DisplayOrder.PadLeft(3, '9');
            else
                xx = x.DisplayOrder.PadLeft(3, '0');

            xx += x.Name;

            string yy = "";
            if (y.DisplayOrder == "")
                yy = y.DisplayOrder.PadLeft(3, '9');
            else
                yy = y.DisplayOrder.PadLeft(3, '0');
            yy += y.Name;

            return xx.CompareTo(yy);
        }
    }
}
