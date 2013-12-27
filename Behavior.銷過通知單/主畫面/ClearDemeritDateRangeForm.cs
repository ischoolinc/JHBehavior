using System;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using K12.Data.Configuration;

namespace JHBehavior.銷過通知單
{
    public partial class ClearDemeritDateRangeForm : SelectDateRangeForm
    {
        private MemoryStream _template = null;
        private byte[] _buffer = null;

        //private bool _preferenceLoaded = false;
        //private bool _printHasRecordOnly = true;

        private string _receiveName = "";
        private string _receiveAddress = "";

        public string ReceiveName { get { return _receiveName; } }
        public string ReceiveAddress { get { return _receiveAddress; } }

        public MemoryStream Template
        {
            get
            {
                MemoryStream _defaultTemplate;

                if (_useDefaultTemplate == "自訂範本")
                {
                    return _template;
                }
                else if (_useDefaultTemplate == "預設範本2")
                {
                    return  _defaultTemplate = new MemoryStream(Properties.Resources.銷過通知單);
                }
                else //預設範本1 
                {
                    return  _defaultTemplate = new MemoryStream(Properties.Resources.銷過通知單_住址中間範本);
                }
            }
        }

        private DateRangeModeNew _mode = DateRangeModeNew.Month;

        private string _useDefaultTemplate = "預設範本1";

        public bool PrintHasRecordOnly
        {
            get { return true; }
        }

        //是否列印學生清單
        private bool _PrintStudentList = false;
        public bool PrintStudentList
        {
            get { return _PrintStudentList; }
        }

        public ClearDemeritDateRangeForm()
        {
            InitializeComponent();
            Text = "銷過通知單";
            LoadPreference();
            InitialDateRange();
        }

        private void LoadPreference()
        {
            #region 讀取 Preference
            ConfigData cd = K12.Data.School.Configuration["銷過通知單_ForK12"];
            XmlElement config = cd.GetXml("XmlData", null);

            if (config != null)
            {
                _useDefaultTemplate = config.GetAttribute("Default");

                XmlElement customize = (XmlElement)config.SelectSingleNode("CustomizeTemplate");
                XmlElement dateRangeMode = (XmlElement)config.SelectSingleNode("DateRangeMode");
                XmlElement receive = (XmlElement)config.SelectSingleNode("Receive");
                XmlElement PrintStudentList = (XmlElement)config.SelectSingleNode("PrintStudentList");

                if (customize != null)
                {
                    string templateBase64 = customize.InnerText;
                    _buffer = Convert.FromBase64String(templateBase64);
                    _template = new MemoryStream(_buffer);
                }

                //列印學生清單
                if (PrintStudentList != null)
                {
                    if (PrintStudentList.HasAttribute("Checked"))
                    {
                        _PrintStudentList = bool.Parse(PrintStudentList.GetAttribute("Checked"));
                    }
                }
                else
                {
                    XmlElement newPrintStudentList = config.OwnerDocument.CreateElement("PrintStudentList");
                    newPrintStudentList.SetAttribute("Checked", "False");
                    config.AppendChild(newPrintStudentList);
                    cd.SetXml("XmlData", config);
                }


                if (receive != null)
                {
                    _receiveName = receive.GetAttribute("Name");
                    _receiveAddress = receive.GetAttribute("Address");
                }
                else
                {
                    XmlElement newReceive = config.OwnerDocument.CreateElement("Receive");
                    newReceive.SetAttribute("Name", "");
                    newReceive.SetAttribute("Address", "");
                    config.AppendChild(newReceive);
                    cd.SetXml("XmlData", config);
                }

                if (dateRangeMode != null)
                {
                    _mode = (DateRangeModeNew)int.Parse(dateRangeMode.InnerText);
                    if (_mode != DateRangeModeNew.Custom)
                        dateTimeInput2.Enabled = false;
                    else
                        dateTimeInput2.Enabled = true;
                }
                else
                {
                    XmlElement newDateRangeMode = config.OwnerDocument.CreateElement("DateRangeMode");
                    newDateRangeMode.InnerText = ((int)_mode).ToString();
                    config.AppendChild(newDateRangeMode);
                    cd.SetXml("XmlData", config);
                }
            }
            else
            {
                #region 產生空白設定檔
                config = new XmlDocument().CreateElement("銷過通知單");
                config.SetAttribute("Default", "預設範本1");
                XmlElement customize = config.OwnerDocument.CreateElement("CustomizeTemplate");
                XmlElement dateRangeMode = config.OwnerDocument.CreateElement("DateRangeMode");
                XmlElement receive = config.OwnerDocument.CreateElement("Receive");
                XmlElement printStudentList = config.OwnerDocument.CreateElement("PrintStudentList");

                dateRangeMode.InnerText = ((int)_mode).ToString();
                receive.SetAttribute("Name", "");
                receive.SetAttribute("Address", "");
                printStudentList.SetAttribute("Checked", "false");

                config.AppendChild(customize);
                config.AppendChild(dateRangeMode);
                config.AppendChild(receive);
                config.AppendChild(printStudentList);

                cd.SetXml("XmlData", config);
                //CurrentUser.Instance.Preference["銷過通知單"] = config;

                _useDefaultTemplate = "預設範本1";
                //_printHasRecordOnly = true;

                #endregion
            }

            cd.Save(); //儲存組態資料。

            #endregion

            //_preferenceLoaded = true;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //傳入True是因為不影響程式結構
            ClearDemeritConfigForm configForm = new ClearDemeritConfigForm(
                _useDefaultTemplate, _mode, _buffer, _receiveName, _receiveAddress, _PrintStudentList);

            if (configForm.ShowDialog() == DialogResult.OK)
            {
                LoadPreference();
                InitialDateRange();
            }
        }

        private void InitialDateRange()
        {
            switch (_mode)
            {
                case DateRangeModeNew.Month: //月
                    {
                        DateTime a = dateTimeInput1.Value;
                        a = GetMonthFirstDay(a);
                        dateTimeInput1.Text = a.ToShortDateString();
                        dateTimeInput2.Text = a.AddMonths(1).AddDays(-1).ToShortDateString();
                        break;
                    }
                case DateRangeModeNew.Week: //週
                    {
                        DateTime b = dateTimeInput1.Value;
                        b = GetWeekFirstDay(b);
                        dateTimeInput1.Text = b.ToShortDateString();
                        dateTimeInput2.Text = b.AddDays(5).ToShortDateString();
                        break;
                    }
                case DateRangeModeNew.Custom: //自訂
                    {
                        //dateTimeInput2.Text = dateTimeInput1.Text = DateTime.Today.ToShortDateString();
                        break;
                    }
                default:
                    throw new Exception("Date Range Mode Error.");
            }

            _printable = true;
            _startTextBoxOK = true;
            _endTextBoxOK = true;
        }

        private DateTime GetWeekFirstDay(DateTime inputDate)
        {
            switch (inputDate.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    return inputDate;
                case DayOfWeek.Tuesday:
                    return inputDate.AddDays(-1);
                case DayOfWeek.Wednesday:
                    return inputDate.AddDays(-2);
                case DayOfWeek.Thursday:
                    return inputDate.AddDays(-3);
                case DayOfWeek.Friday:
                    return inputDate.AddDays(-4);
                case DayOfWeek.Saturday:
                    return inputDate.AddDays(-5);
                default:
                    return inputDate.AddDays(-6);
            }
        }

        private DateTime GetMonthFirstDay(DateTime inputDate)
        {
            return DateTime.Parse(inputDate.Year + "/" + inputDate.Month + "/1");
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //if (_printable)
            //    dateTimeInput1.Text = _startDate.ToShortDateString();
            //timer1.Stop();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dateTimeInput1_TextChanged(object sender, EventArgs e)
        {
            if (_startTextBoxOK && _mode != DateRangeModeNew.Custom)
            {
                switch (_mode)
                {
                    case DateRangeModeNew.Month: //月
                        {
                            _startDate = GetMonthFirstDay(DateTime.Parse(dateTimeInput1.Text));
                            _endDate = _startDate.AddMonths(1).AddDays(-1);
                            dateTimeInput1.Text = _startDate.ToShortDateString();
                            dateTimeInput2.Text = _endDate.ToShortDateString();
                            _printable = true;
                            break;
                        }
                    case DateRangeModeNew.Week: //週
                        {
                            _startDate = GetWeekFirstDay(DateTime.Parse(dateTimeInput1.Text));
                            _endDate = _startDate.AddDays(4);
                            dateTimeInput1.Text = _startDate.ToShortDateString();
                            dateTimeInput2.Text = _endDate.ToShortDateString();
                            _printable = true;
                            break;
                        }
                    case DateRangeModeNew.Custom: //自訂
                        break;
                    default:
                        throw new Exception("Date Range Mode Error");
                }

                //if (dateTimeInput1.Text != _startDate.ToShortDateString() && timer1 != null)
                //    timer1.Start();
                errorProvider1.Clear();
            }
        }

        private void dateTimeInput2_TextChanged(object sender, EventArgs e)
        {
            //if (_preferenceLoaded)
            //{
            //    if (_mode == DateRangeModeNew.Custom)
            //    {
            //        base.dateTimeInput2_TextChanged(sender, e);
            //    }
            //    else
            //    {
            //        _endTextBoxOK = true;
            //        errorProvider2.Clear();
            //    }
            //}
        }
    }

    public enum DateRangeModeNew { Month, Week, Custom }
}