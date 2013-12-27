using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Framework;
using System.Xml;
using FISCA.DSAUtil;
using System.IO;

namespace Behavior.ClubActivitiesPointList
{
    public class GetConfigData
    {
        string _ConfigName;
        K12.Data.Configuration.ConfigData cd;

        public XmlElement BeforeSetup;
        public XmlElement PrintSetup;
        public XmlElement WeekSetup;

        //string _FormCheckMode;
        //int _WeekONE;
        //int _WeekTWO;
        public string _Day1;
        public string _Day2;

        public string _PrintMode;
        public string _PrintTemp;
        public byte[] _buffer;
        public int _Week;
        public MemoryStream Setup_template = null;

        //public string CheckBool;

        public GetConfigData(string ConfigName)
        {
            _ConfigName = ConfigName;

            cd = K12.Data.School.Configuration[_ConfigName];

            if (cd.Count == 0) //如果沒有任何設定內容
            {
                //前次輸入狀態
                DSXmlHelper dsx1 = new DSXmlHelper("BeforeSetup");
                dsx1.AddElement("Day1");
                dsx1.SetAttribute("Day1", "day", DateTime.Now.ToShortDateString());
                dsx1.AddElement("Day2");
                dsx1.SetAttribute("Day2", "day", DateTime.Now.ToShortDateString());
                cd["前次輸入狀態"] = dsx1.BaseElement.OuterXml;
                
                //列印設定
                DSXmlHelper dsx2 = new DSXmlHelper("PrintSetup");
                dsx2.AddElement("PrintMode");
                dsx2.SetAttribute("PrintMode", "bool", "true"); //true為使用預設
                dsx2.AddElement("PrintTemp");
                dsx2.SetAttribute("PrintTemp", "Temp", ""); //true為使用預設
                dsx2.AddElement("Week");
                dsx2.SetAttribute("Week", "day", "0"); //預設選星期一
                cd["列印設定"] = dsx2.BaseElement.OuterXml;

                //DSXmlHelper dsx3 = new DSXmlHelper("WeekSetup");
                //dsx3.AddElement("OneWeekDay");
                //dsx3.SetAttribute("OneWeekDay", "Day", DateTime.Now.ToShortDateString());
                //cd["週次設定"] = dsx3.BaseElement.OuterXml;

            }

            Reset();
        }

        public void Reset()
        {
            
            BeforeSetup = DSXmlHelper.LoadXml(cd["前次輸入狀態"]);
            _Day1 = (BeforeSetup.SelectSingleNode("Day1") as XmlElement).GetAttribute("day");
            _Day2 = (BeforeSetup.SelectSingleNode("Day2") as XmlElement).GetAttribute("day");

            PrintSetup = DSXmlHelper.LoadXml(cd["列印設定"]);
            _PrintMode = (PrintSetup.SelectSingleNode("PrintMode") as XmlElement).GetAttribute("bool");
            _PrintTemp = (PrintSetup.SelectSingleNode("PrintTemp") as XmlElement).GetAttribute("Temp");
            if (_PrintTemp != "")
            {
                _buffer = Convert.FromBase64String(_PrintTemp);
                Setup_template = new MemoryStream(_buffer);
            }
            _Week = int.Parse((PrintSetup.SelectSingleNode("Week") as XmlElement).GetAttribute("day"));

            //_FormCheckMode = (BeforeSetup.SelectSingleNode("Check") as XmlElement).GetAttribute("bool");
            //_WeekONE = int.Parse((BeforeSetup.SelectSingleNode("Week1") as XmlElement).GetAttribute("int"));
            //_WeekTWO = int.Parse((BeforeSetup.SelectSingleNode("Week2") as XmlElement).GetAttribute("int"));
            //WeekSetup = DSXmlHelper.LoadXml(cd["週次設定"]);
        }


        public void SaveBefore(XmlElement before)
        {
            cd["前次輸入狀態"] = before.OuterXml;
            SaveAll();
        }

        public void SavePrint(XmlElement PrintText)
        {
            cd["列印設定"] = PrintText.OuterXml;
            SaveAll();
        }

        //public void SaveWeekDay(string day)
        //{
        //    cd["週次設定"] = day;
        //    SaveAll();
        //}

        private void SaveAll()
        {
            cd.Save();
            Reset();
        }


    }
}
