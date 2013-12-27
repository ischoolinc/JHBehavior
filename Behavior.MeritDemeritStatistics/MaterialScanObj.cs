using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;
using System.IO;
using JHSchool.Data;
using K12.Logic;

namespace Behavior.MeritDemeritStatistics
{
    class MaterialScanObj
    {
        /// <summary>
        /// 男女分類後的自動統計資料
        /// </summary>
        AcquisitionOfInformation AOI;

        /// <summary>
        /// 統計相關獎懲資料
        /// </summary>
        MeritDemeritObj MDObj = new MeritDemeritObj();

        Workbook book;

        /// <summary>
        /// 傳入資料清單,即可進行掃瞄
        /// </summary>
        /// <param name="_AOI"></param>
        public MaterialScanObj(AcquisitionOfInformation _AOI)
        {
            AOI = _AOI;

            Scan();
        }

        /// <summary>
        /// 資料掃瞄開始
        /// </summary>
        private void Scan()
        {
            MDObj.一年級男生獎懲 = AutoSummaryScan(AOI.FirstGradeMale);
            MDObj.二年級男生獎懲 = AutoSummaryScan(AOI.SecondGradeMale);
            MDObj.三年級男生獎懲 = AutoSummaryScan(AOI.ThirdGradeMale);

            MDObj.一年級女生獎懲 = AutoSummaryScan(AOI.FirstGradeFemale);
            MDObj.二年級女生獎懲 = AutoSummaryScan(AOI.SecondGradeFemale);
            MDObj.三年級女生獎懲 = AutoSummaryScan(AOI.ThirdGradeFemale);

            MDObj.男生獎懲 = AutoSummaryScan(AOI.Male);
            MDObj.女生獎懲 = AutoSummaryScan(AOI.Female);

            MDObj.總獎懲 = AutoSummaryScan(AOI.ALL);
        }

        /// <summary>
        /// 傳入AutoSummary清單,以取得獎懲統計狀況
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        private Dictionary<string,int> AutoSummaryScan(List<AutoSummaryRecord> list)
        {
            //以學生ID為Key,每名學生有List筆記錄
            Dictionary<string, List<AutoSummaryRecord>> test = new Dictionary<string, List<AutoSummaryRecord>>();

            foreach (AutoSummaryRecord each in list)
            {
                if (!test.ContainsKey(each.RefStudentID))
                {
                    test.Add(each.RefStudentID, new List<AutoSummaryRecord>());
                }
                test[each.RefStudentID].Add(each);
            }



            Dictionary<string, int> dic = new Dictionary<string, int>();
            dic.Add("獎勵總人數", 0);
            dic.Add("嘉獎人數", 0);
            dic.Add("嘉獎次數", 0);
            dic.Add("小功人數", 0);
            dic.Add("小功次數", 0);
            dic.Add("大功人數", 0);
            dic.Add("大功次數", 0);

            dic.Add("懲戒總人數", 0);
            dic.Add("警告人數", 0);
            dic.Add("警告次數", 0);
            dic.Add("小過人數", 0);
            dic.Add("小過次數", 0);
            dic.Add("大過人數", 0);
            dic.Add("大過次數", 0);

            dic.Add("銷過總人數", 0);
            dic.Add("警告銷過人數", 0);
            dic.Add("警告銷過次數", 0);
            dic.Add("小過銷過人數", 0);
            dic.Add("小過銷過次數", 0);
            dic.Add("大過銷過人數", 0);
            dic.Add("大過銷過次數", 0);

            foreach (string studentId in test.Keys)
            {
                //用來避免重覆記錄人次的問題
                bool studentLoop1 = true;
                bool studentLoop2 = true;
                bool studentLoop3 = true;
                bool studentLoop4 = true;
                bool studentLoop5 = true;
                bool studentLoop6 = true;
                bool studentLoop7 = true;
                bool studentLoop8 = true;
                bool studentLoop9 = true;
                bool studentLoop10 = true;
                bool studentLoop11 = true;
                bool studentLoop12 = true;

                foreach (AutoSummaryRecord auto in test[studentId])
                {
                    #region 獎勵
                    if (auto.MeritA + auto.MeritB + auto.MeritC > 0)
                    {
                        if (studentLoop1)
                        {
                            dic["獎勵總人數"]++;
                            studentLoop1 = false;
                        }

                         if (auto.MeritC > 0)
                        {
                            if (studentLoop2)
                            {
                                dic["嘉獎人數"]++;
                                studentLoop2 = false;
                            }
                            dic["嘉獎次數"] += auto.MeritC;
                        }
                        if (auto.MeritB > 0)
                        {
                            if (studentLoop3)
                            {
                                dic["小功人數"]++;
                                studentLoop3 = false;
                            }
                            dic["小功次數"] += auto.MeritB;
                        }
                        if (auto.MeritA > 0)
                        {
                            if (studentLoop4)
                            {
                                dic["大功人數"]++;
                                studentLoop4 = false;
                            }
                            dic["大功次數"] += auto.MeritA;
                        }
                    }
                    #endregion

                    #region 懲戒

                    if (auto.DemeritA + auto.DemeritB + auto.DemeritC > 0)
                    {
                        if (studentLoop5)
                        {
                            dic["懲戒總人數"]++;
                            studentLoop5 = false;
                        }

                        if (auto.DemeritC > 0)
                        {
                            if (studentLoop6)
                            {
                                dic["警告人數"]++;
                                studentLoop6 = false;
                            }
                            dic["警告次數"] += auto.DemeritC;
                        }
                        if (auto.DemeritB > 0)
                        {
                            if (studentLoop7)
                            {
                                dic["小過人數"]++;
                                studentLoop7 = false;
                            }
                            dic["小過次數"] += auto.DemeritB;
                        }
                        if (auto.DemeritA > 0)
                        {
                            if (studentLoop8)
                            {
                                dic["大過人數"]++;
                                studentLoop8 = false;
                            }
                            dic["大過次數"] += auto.DemeritA;
                        }
                    }
                    #endregion

                    #region 銷過

                    if (auto.ClearedDemeritA + auto.ClearedDemeritB + auto.ClearedDemeritC > 0)
                    {
                        if (studentLoop9)
                        {
                            dic["銷過總人數"]++;
                            studentLoop9 = false;
                        }

                        if (auto.ClearedDemeritC > 0)
                        {
                            if (studentLoop10)
                            {
                                dic["警告銷過人數"]++;
                                studentLoop10 = false;
                            }
                            dic["警告銷過次數"] += auto.ClearedDemeritC;
                        }
                        if (auto.ClearedDemeritB > 0)
                        {
                            if (studentLoop11)
                            {
                                dic["小過銷過人數"]++;
                                studentLoop11 = false;
                            }
                            dic["小過銷過次數"] += auto.ClearedDemeritB;
                        }
                        if (auto.ClearedDemeritA > 0)
                        {
                            if (studentLoop12)
                            {
                                dic["大過銷過人數"]++;
                                studentLoop12 = false;
                            }
                            dic["大過銷過次數"] += auto.ClearedDemeritA;
                        }
                    }
                    #endregion
                }
            }

            return dic;
        }

        /// <summary>
        /// 將掃瞄後資料建立為Excel並且傳出
        /// </summary>
        /// <returns></returns>
        public Workbook CreateExcel()
        {
            book = new Workbook();
            book.Open(new MemoryStream(Properties.Resources.範本), FileFormatType.Excel2003);

            SetValue(MDObj.總獎懲, 2);
            SetValue(MDObj.男生獎懲, 3);
            SetValue(MDObj.女生獎懲, 4);
            SetValue(MDObj.一年級男生獎懲, 5);
            SetValue(MDObj.一年級女生獎懲, 6);
            SetValue(MDObj.二年級男生獎懲, 7);
            SetValue(MDObj.二年級女生獎懲, 8);
            SetValue(MDObj.三年級男生獎懲, 9);
            SetValue(MDObj.三年級女生獎懲, 10);

            return book;
        }

        private void SetValue(Dictionary<string, int> dic, int x)
        {
            int index = 2;
            foreach (int each in dic.Values)
            {
                book.Worksheets[0].Cells[index, x].PutValue(each);
                if (index == 8)
                {
                    index++;
                }
                if (index == 16)
                {
                    index++;
                }
                index++;
            }            
        }
    }
}
