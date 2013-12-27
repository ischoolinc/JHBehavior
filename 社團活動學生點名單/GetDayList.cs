using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Presentation.Controls;

namespace Behavior.ClubActivitiesPointList
{
    class GetDayList
    {
        private DateTime _StartDay;
        private DateTime _EndDay;

        /// <summary>
        /// 指定日期區間內,特定星期的清單
        /// </summary>
        public List<DateTime> DateTimeList = new List<DateTime>();

        /// <summary>
        /// 依 星期X 編號,取得該區間內有多少 星期X
        /// </summary>
        /// <param name="Day">星期幾的代號</param>
        /// <param name="StartDay">開始日期</param>
        /// <param name="EndDay">結束日期</param>
        public GetDayList(string Day,string StartDay,string EndDay)
        {
            #region 依 星期X 編號,取得該區間內有多少 星期X
            if (CheckDay(StartDay) && CheckDay(EndDay))
            {
                _StartDay = DateTime.Parse(StartDay);
                _EndDay = DateTime.Parse(EndDay);
            }
            else
            {
                MsgBox.Show("輸入時間有誤!");
                return;
            }

            //取得總日數
            TimeSpan ts = _EndDay - _StartDay;
            List<DateTime> list = new List<DateTime>();
            for (int x = 0; x <= ts.Days; x++)
            {
                list.Add(_StartDay.AddDays(x));
            }

            //取得特定日期
            DayOfWeek DayWeek = CheckWeek(Day);
            foreach (DateTime each in list)
            {
                if (each.DayOfWeek == DayWeek)
                {
                    DateTimeList.Add(each);
                }
            } 
            #endregion
        }

        private DayOfWeek CheckWeek(string x)
        {
            #region 依編號取代為星期
            if (x == "0")
            {
                return DayOfWeek.Monday;
            }
            else if (x == "1")
            {
                return DayOfWeek.Tuesday;
            }
            else if (x == "2")
            {
                return DayOfWeek.Wednesday;
            }
            else if (x == "3")
            {
                return DayOfWeek.Thursday;
            }
            else if (x == "4")
            {
                return DayOfWeek.Friday;
            }
            else if (x == "5")
            {
                return DayOfWeek.Saturday;
            }
            else
            {
                return DayOfWeek.Sunday;
            } 
            #endregion
        }

        private bool CheckDay(string now)
        {
            #region 日期判斷式
            DateTime StartDate;
            if (DateTime.TryParse(now, out StartDate))
            {
                return true;
            }
            else
            {
                return false;
            } 
            #endregion
        }
    }
}
