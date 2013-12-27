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
namespace JHSchool.Behavior.StuAdminExtendControls
{
    public partial class SchoolDayForm : FISCA.Presentation.Controls.BaseForm
    {
        private K12.Data.SchoolHolidayRecord record;
        private List<DevComponents.AdvTree.Node> allDays;

        public SchoolDayForm()
        {
            InitializeComponent();
            
        }

        //1。找出目前學年度學期
        //2. 查詢是否已有設定值，若有，就讀出來。
        //3. 若無設定值，就設定為今天。
        private void PeriodInSchool_Load(object sender, EventArgs e)
        {
            schoolyear.Text = School.DefaultSchoolYear + "學年度第" + School.DefaultSemester + "學期";

            record = K12.Data.SchoolHoliday.SelectSchoolHolidayRecord();

            if (record == null)
            {
                this.dtBeginDate.Value = DateTime.Today;
                this.dtEndDate.Value = DateTime.Today;  //會觸發 FillDateList()
            }
            else
            {
                this.dtBeginDate.Value = this.record.BeginDate;
                this.dtEndDate.Value = this.record.EndDate;

                this.txtG1Count.Text = record.SchoolDayCountG1.ToString();
                this.txtG2Count.Text = record.SchoolDayCountG2.ToString();
                this.txtG3Count.Text = record.SchoolDayCountG3.ToString();
            }
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

            this.txtG1Count.Text = schoolDays.ToString();
            this.txtG2Count.Text = schoolDays.ToString();
            this.txtG3Count.Text = schoolDays.ToString();
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

            result = (dt.DayOfWeek == DayOfWeek.Saturday || dt.DayOfWeek == DayOfWeek.Sunday);

            if (this.record != null)
            {
                if (this.record.IsContained(dt))
                    result =  this.record.IsHoliday(dt);
            }

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
        private bool txtCheck(string txt,TextBoxX Text)
        {  
            uint x = 0;
            if (uint.TryParse(txt, out  x) )
            {
                errorProvider1.Clear();
                return true;
            }
            else
            {
                errorProvider1.SetError(Text, "請輸入正確的數字");
                
                return false;
            }
        }


        private void SetButton_Click(object sender, EventArgs e)
        {
            if (this.record == null)
                this.record = new K12.Data.SchoolHolidayRecord();

            this.record.BeginDate = this.dtBeginDate.Value;
            this.record.EndDate = this.dtEndDate.Value;

            if (txtCheck(this.txtG1Count.Text, this.txtG1Count))
                this.record.SchoolDayCountG1 =int.Parse(this.txtG1Count.Text);
            else return;

            if(txtCheck(this.txtG2Count.Text,this.txtG2Count))
                this.record.SchoolDayCountG2 = int.Parse(this.txtG2Count.Text);
            else return;

            if(txtCheck(this.txtG3Count.Text,this.txtG3Count))
                this.record.SchoolDayCountG3 = int.Parse(this.txtG3Count.Text);
            else return;

            this.record.HolidayList.Clear();

            if (this.allDays == null)
            {
                FISCA.Presentation.Controls.MsgBox.Show("請設定學期開始及學期結束日期");
                return;
            }

            foreach (DevComponents.AdvTree.Node nd in this.allDays)
            {
                if (nd.Cells[1].Checked)
                    this.record.HolidayList.Add((DateTime)nd.Tag);
            }

            K12.Data.SchoolHoliday.SetSchoolHolidayRecord(this.record);

            this.Close();

        }
      

        
    }
}
