using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SmartSchool.Common;
using Framework;
using DevComponents.DotNetBar.Controls;
using K12.Data;
using FISCA.Data;
using System.Xml.Linq;
using System.Xml.XPath;

namespace JHSchool.Behavior.StuAdminExtendControls
{
    public partial class SchoolDayForm : FISCA.Presentation.Controls.BaseForm
    {
        //因為K12只提供到SchoolDayCountG3,所以改為直接存取
        //private K12.Data.SchoolHolidayRecord record;
        private List<DevComponents.AdvTree.Node> allDays;

        private const string SchoolHodidayConfigString = "SCHOOL_HOLIDAY_CONFIG_STRING";
        private const string configString = "CONFIG_STRING";

        private QueryHelper _Q;
        private List<DateTime> _Holidays;
        
        private ConfigData _CD;
        private DateTime _OldStartDate;
        private DateTime _OldEndDate;

        public SchoolDayForm()
        {
            InitializeComponent();
            _Q = new QueryHelper();
            _Holidays = new List<DateTime>();
            _CD = School.Configuration[SchoolHodidayConfigString];
        }

        //1。找出目前學年度學期
        //2. 查詢是否已有設定值，若有，就讀出來。
        //3. 若無設定值，就設定為今天。
        private void PeriodInSchool_Load(object sender, EventArgs e)
        {
            List<string> cols = new List<string>() { "年級", "上課天數" };
            Campus.Windows.DataGridViewImeDecorator dec = new Campus.Windows.DataGridViewImeDecorator(this.dgv, cols);


            schoolyear.Text = School.DefaultSchoolYear + "學年度第" + School.DefaultSemester + "學期";

            //取得之前設定設定
            XElement rootXml = null;
            string xmlContent = _CD[configString];

            if (!string.IsNullOrWhiteSpace(xmlContent))
                rootXml = XElement.Parse(xmlContent);
            else
                rootXml = new XElement("SchoolHolidays");

            //讀出之前的假日清單
            foreach (XElement holiday in rootXml.XPathSelectElements("//Holiday"))
            {
                DateTime date;
                if (DateTime.TryParse(holiday.Value, out date))
                    _Holidays.Add(date);
            }

            //開始時間
            XElement BeginDate = rootXml.XPathSelectElement("//BeginDate");
            _OldStartDate = BeginDate == null || string.IsNullOrWhiteSpace(BeginDate.Value) ? DateTime.Today : DateTime.Parse(BeginDate.Value);
            this.dtBeginDate.Value = _OldStartDate;

            //結束時間
            XElement EndDate = rootXml.XPathSelectElement("//EndDate");
            _OldEndDate = EndDate == null || string.IsNullOrWhiteSpace(EndDate.Value) ? DateTime.Today : DateTime.Parse(EndDate.Value);
            this.dtEndDate.Value = _OldEndDate;

            //取得全校既有的年級並帶入設定
            DataTable dt = _Q.Select("select distinct grade_year from class where grade_year is not null order by grade_year");
            foreach (DataRow row in dt.Rows)
            {
                DataGridViewRow dgvRow = new DataGridViewRow();
                dgvRow.Tag = row["grade_year"] + "";
                string grade = "//SchoolDayCountG" + dgvRow.Tag;
                XElement elem = rootXml.XPathSelectElement(grade);
                string value = elem == null ? string.Empty : elem.Value;
                dgvRow.CreateCells(dgv, dgvRow.Tag + "年級", value);
                dgv.Rows.Add(dgvRow);
            }

            //record = K12.Data.SchoolHoliday.SelectSchoolHolidayRecord();

            //if (record == null)
            //{
            //    this.dtBeginDate.Value = DateTime.Today;
            //    this.dtEndDate.Value = DateTime.Today;  //會觸發 FillDateList()
            //}
            //else
            //{
            //    this.dtBeginDate.Value = this.record.BeginDate;
            //    this.dtEndDate.Value = this.record.EndDate;

            //    this.txtG1Count.Text = record.SchoolDayCountG1.ToString();
            //    this.txtG2Count.Text = record.SchoolDayCountG2.ToString();
            //    this.txtG3Count.Text = record.SchoolDayCountG3.ToString();
            //}
        }


        /// <summary>
        /// 根據開始和結束日期，找出此區間內所有日期，並填入AdvTree中。
        /// </summary>
        private void FillDateList()
        {
            if (dtEndDate.Value <= this.dtBeginDate.Value)
                return;

            this.allDays = new List<DevComponents.AdvTree.Node>();  //紀錄畫面上所有的日期清單

            int year = 0;
            int month = 0;
            int day = 0;

            DevComponents.AdvTree.Node ndYear= new DevComponents.AdvTree.Node();
            DevComponents.AdvTree.Node ndMonth = new DevComponents.AdvTree.Node();
            DevComponents.AdvTree.Node ndDay = new DevComponents.AdvTree.Node();

            this.advTree1.Nodes.Clear();

            DateTime currentDate = this.dtBeginDate.Value;

            while (this.dtEndDate.Value >= currentDate)
            {
                if (currentDate.Year != year)   //不同年
                {
                    ndYear = createYearNode(currentDate);                                    
                    this.advTree1.Nodes.Add(ndYear);
                    year = currentDate.Year;

                    ndMonth = createMonthNode(currentDate);
                    ndYear.Nodes.Add(ndMonth);
                    month = currentDate.Month;

                    ndDay = createDayNode(currentDate);                    
                    ndMonth.Nodes.Add(ndDay);
                    day = currentDate.Day;
                }
                else if (currentDate.Month != month)    //同年不同月
                {
                    ndMonth = createMonthNode(currentDate);
                    ndYear.Nodes.Add(ndMonth);
                    month = currentDate.Month;

                    ndDay = createDayNode(currentDate);
                    ndMonth.Nodes.Add(ndDay);
                    day = currentDate.Day;
                }
                else    //同年同月不同日
                {
                    ndDay = createDayNode(currentDate);
                    ndMonth.Nodes.Add(ndDay);
                    day = currentDate.Day;
                }

                this.allDays.Add(ndDay);

                currentDate = currentDate.AddDays(1);
            }

            this.calculateSchoolDays();
        }

        private DevComponents.AdvTree.Node createYearNode(DateTime currentDate)
        {
            DevComponents.AdvTree.Node ndYear = new DevComponents.AdvTree.Node();
            ndYear.Cells[0].Text = currentDate.Year.ToString() + " 年";
            return ndYear;
        }

        private DevComponents.AdvTree.Node createMonthNode(DateTime currentDate)
        {
            DevComponents.AdvTree.Node ndMonth = new DevComponents.AdvTree.Node();
            ndMonth.Cells[0].Text = currentDate.Month.ToString() + " 月";
            return ndMonth;
        }

        private DevComponents.AdvTree.Node createDayNode(DateTime currentDate)
        {
            DevComponents.AdvTree.Node ndDay = new DevComponents.AdvTree.Node();
            
            ndDay.Cells[0].Text = toChineseWeekday(currentDate);
            ndDay.Cells.Add(createCell(currentDate));
            ndDay.Tag = currentDate;
            ndDay.NodeClick += new EventHandler(ndDay_NodeClick);

            return ndDay;
        }
            
        void ndDay_NodeClick(object sender, EventArgs e)
        {
            this.calculateSchoolDays();
        }

        //要重算上課天數
        private void calculateSchoolDays()
        {
            int schoolDays = 0;
            foreach (DevComponents.AdvTree.Node nd in this.allDays)
            {
                if (!nd.Cells[1].Checked)
                    schoolDays += 1;
            }

            foreach (DataGridViewRow row in dgv.Rows)
                row.Cells[colDays.Index].Value = schoolDays + "";

            //this.txtG1Count.Text = schoolDays.ToString();
            //this.txtG2Count.Text = schoolDays.ToString();
            //this.txtG3Count.Text = schoolDays.ToString();
        }

        private DevComponents.AdvTree.Cell createCell(DateTime currentDate)
        {
            DevComponents.AdvTree.Cell cell = new DevComponents.AdvTree.Cell();
            cell.CheckBoxStyle = DevComponents.DotNetBar.eCheckBoxStyle.RadioButton;
            cell.CheckBoxVisible = true;
            cell.Checked = isHoliday(currentDate);
            
            return cell;
        }

        private string toChineseWeekday(DateTime dt)
        {
            string result = "";
            switch (dt.DayOfWeek)
            {
                case DayOfWeek.Sunday :
                    result = "日";
                    break;
                case DayOfWeek.Monday:
                    result = "一";
                    break;
                case DayOfWeek.Tuesday:
                    result = "二";
                    break;
                case DayOfWeek.Wednesday:
                    result = "三";
                    break;
                case DayOfWeek.Thursday:
                    result = "四";
                    break;
                case DayOfWeek.Friday:
                    result = "五";
                    break;
                case DayOfWeek.Saturday:
                    result = "六";
                    break;
            }

            result = string.Format("{0} ({1})", dt.ToShortDateString() ,result);

            return result;
        }


        private bool isHoliday(DateTime dt)
        {
            bool result = false;

            //預設六日先打勾
            result = (dt.DayOfWeek == DayOfWeek.Saturday || dt.DayOfWeek == DayOfWeek.Sunday);

            //若該日期在以前設定的日期區間,則以假日清單為準
            if (dt >= _OldStartDate && dt <= _OldEndDate)
            {
                if (_Holidays.Contains(dt))
                    result = true;
                else
                    result = false;
            }

            //if (this.record != null)
            //{
            //    if (this.record.IsContained(dt))
            //        result =  this.record.IsHoliday(dt);
            //}

            return result ;
        }

        private void btn_Exit(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dtEndDate_ValueChanged(object sender, EventArgs e)
        {
            FillDateList();
        }

        private void dtBeginDate_ValueChanged(object sender, EventArgs e)
        {
            FillDateList();
        }

        //private bool txtCheck(string txt,TextBoxX Text)
        //{  
        //    uint x = 0;
        //    if (uint.TryParse(txt, out  x) )
        //    {
        //        errorProvider1.Clear();
        //        return true;
        //    }
        //    else
        //    {
        //        errorProvider1.SetError(Text, "請輸入正確的數字");
                
        //        return false;
        //    }
        //}

        private void SetButton_Click(object sender, EventArgs e)
        {
            bool hasError = false;

            //主XML
            XElement elem = new XElement("SchoolHolidays", new XElement("BeginDate", this.dtBeginDate.Value.ToShortDateString()), new XElement("EndDate", this.dtEndDate.Value.ToShortDateString()));

            //掛上各年級的上課天數節點(順便驗證欄位資料)
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.ErrorText = string.Empty;
                int count;
                if (int.TryParse(row.Cells[colDays.Index].Value + "", out count))
                    elem.Add(new XElement("SchoolDayCountG" + row.Tag, count));
                else
                {
                    row.ErrorText = "上課天數欄位必須輸入數字";
                    hasError = true;
                }
            }

            if (hasError)
            {
                MessageBox.Show("資料有誤,請確認後再儲存...");
                return;
            }
                
            //if (this.record == null)
            //    this.record = new K12.Data.SchoolHolidayRecord();

            //this.record.BeginDate = this.dtBeginDate.Value;
            //this.record.EndDate = this.dtEndDate.Value;

            //if (txtCheck(this.txtG1Count.Text, this.txtG1Count))
            //    this.record.SchoolDayCountG1 = int.Parse(this.txtG1Count.Text);
            //else return;

            //if (txtCheck(this.txtG2Count.Text, this.txtG2Count))
            //    this.record.SchoolDayCountG2 = int.Parse(this.txtG2Count.Text);
            //else return;

            //if (txtCheck(this.txtG3Count.Text, this.txtG3Count))
            //    this.record.SchoolDayCountG3 = int.Parse(this.txtG3Count.Text);
            //else return;

            //this.record.HolidayList.Clear();

            if (this.allDays == null)
            {
                FISCA.Presentation.Controls.MsgBox.Show("請設定學期開始及學期結束日期");
                return;
            }

            //建立畫面設定的假日節點
            XElement holidayList = new XElement("HolidayList");
            foreach (DevComponents.AdvTree.Node nd in this.allDays)
            {
                if (nd.Cells[1].Checked)
                {
                    DateTime dt = (DateTime)nd.Tag;
                    holidayList.Add(new XElement("Holiday", dt.ToShortDateString()));
                }
            }
            //掛上假日節點
            elem.Add(holidayList);

            //foreach (DevComponents.AdvTree.Node nd in this.allDays)
            //{
            //    if (nd.Cells[1].Checked)
            //        this.record.HolidayList.Add((DateTime)nd.Tag);
            //}

            //K12.Data.SchoolHoliday.SetSchoolHolidayRecord(this.record);

            //回存
            _CD[configString] = elem.ToString(SaveOptions.DisableFormatting);
            _CD.Save();

            this.Close();
        }

        private void dgv_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == colDays.Index)
                dgv.BeginEdit(true);
        }
    }
}
