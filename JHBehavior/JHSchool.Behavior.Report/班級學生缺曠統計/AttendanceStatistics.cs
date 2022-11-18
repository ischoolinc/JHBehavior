using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using Aspose.Cells;
using JHSchool.Behavior.Report;
using K12.Data.Configuration;
using FISCA.Presentation.Controls;
using K12.Data;
using System.Drawing;

namespace _71103_classTDK
{
    public partial class AttendanceStatistics : BaseForm
    {
        private ConfigData cd;

        public AttendanceStatistics()
        {
            InitializeComponent();

            #region 處理假別

            cd = K12.Data.School.Configuration["班級出缺席狀況統計表"];
            textBoxX3.Text = cd["每日節次"];

            List<AbsenceMappingInfo> list = AbsenceMapping.SelectAll();

            listView1.Items.Clear();
            foreach (AbsenceMappingInfo absence in list)
            {
                listView1.Items.Add(new ListViewItem(absence.Name));
            }

            for (int x = 0; x < listView1.Items.Count; x++)
            {
                listView1.Items[x].Checked = true;
            }

            dateTimeInput1.Value = DateTime.Today.AddDays(-1);
            dateTimeInput2.Value = DateTime.Today;

            #endregion

        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            this.Enabled = false;

            #region 處理節次

            int gtRoden;
            if (!int.TryParse(textBoxX3.Text, out gtRoden))
            {
                MessageBox.Show("節次內容非數字型態!");
                this.Enabled = true;
                return;
            }

            cd["每日節次"] = textBoxX3.Text;
            cd.Save();

            #endregion

            #region Excel建立

            Workbook classCell = new Workbook(new MemoryStream(ProjectResource.班級學生缺曠統計),new LoadOptions(LoadFormat.Excel97To2003));

            //儲存元件
            SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

            List<ClassRecord> classList = K12.Data.Class.SelectByIDs(K12.Presentation.NLDPanels.Class.SelectedSource);

            classList = SortClassIndex.K12Data_ClassRecord(classList);

            #endregion

            #region 建立字典&範本

            Dictionary<string, int> AttList = new Dictionary<string, int>();
            int test1 = 4;
            foreach (ListViewItem var in listView1.Items)
            {
                if (var.Checked == true)
                {
                    AttList.Add(var.Text, 0);
                    classCell.Worksheets[0].Cells[1, test1].PutValue(var.Text);
                    test1++;
                }
            }

            #endregion

            #region 處理日期

            if (dateTimeInput1.Value > dateTimeInput2.Value)
            {
                MessageBox.Show("開始時間大於結束時間!!");
                return;
            }


            int dayNum = 0; //取得日期天數
            DateTime Temp = dateTimeInput1.Value;
            while (Temp <= dateTimeInput2.Value)
            {
                if (Temp.DayOfWeek == DayOfWeek.Saturday || Temp.DayOfWeek == DayOfWeek.Sunday) //六/日不列入出席率計算
                {
                }
                else
                {
                    dayNum++;
                }
                Temp = Temp.AddDays(1);
            }

            #endregion


            int nX = 2;
            double TaClass = 0;
            double Att = 0;

            foreach (ClassRecord classRecord in classList)
            {
                List<StudentRecord> studentList = new List<StudentRecord>();
                foreach (StudentRecord student in classRecord.Students)
                {
                    if (student.Status == StudentRecord.StudentStatus.一般 || student.Status == StudentRecord.StudentStatus.輟學)
                    {
                        studentList.Add(student);
                    }
                }
                //classCell.Worksheets[0].Cells[nX, 0].PutValue(f.Department);
                classCell.Worksheets[0].Cells[nX, 1].PutValue(classRecord.Name);

                classCell.Worksheets[0].Cells[nX, 2].PutValue(studentList.Count);


                if (classRecord.Teacher != null)
                {
                    classCell.Worksheets[0].Cells[nX, 3].PutValue(classRecord.Teacher.Name);
                }

                List<AttendanceRecord> AttendanceList = Attendance.SelectByDate(studentList, dateTimeInput1.Value, dateTimeInput2.Value);

                #region 處理缺曠內容


                foreach (AttendanceRecord AttInfo in AttendanceList)
                {
                    foreach (AttendancePeriod AttPer in AttInfo.PeriodDetail)
                    {
                        //dateTimeInput2.Value.AddDays(1); //多一天??

                        if (AttList.ContainsKey(AttPer.AbsenceType))
                        {
                            AttList[AttPer.AbsenceType]++;
                        }
                    }
                }
                #endregion


                int DicNum = 4;
                //
                foreach (string var in AttList.Keys)
                {
                    classCell.Worksheets[0].Cells[nX, DicNum].PutValue(AttList[var]);
                    DicNum++;
                    Att = Att + AttList[var];
                }

                for (int x = 0; x < AttList.Count; x++)
                {
                    AttList[listView1.Items[x].Text] = 0;
                }

                //計算出席的率公式

                //Att                           總缺礦數
                //classRecord.Students.Count    人數
                //gtRoden                       一天的節數
                //dayNum                        非假日天數
                if (studentList.Count != 0)
                {
                    TaClass = 1 - ((double)Att / (studentList.Count * gtRoden * dayNum));
                }
                TaClass = TaClass * 100;
                classCell.Worksheets[0].Cells[nX, DicNum].PutValue(Math.Round(TaClass, 2, MidpointRounding.AwayFromZero) + "%");
                classCell.Worksheets[0].Cells[1, DicNum].PutValue("出席率");


                TaClass = 0;
                Att = 0;
                nX++;

            }

            this.Enabled = true;

            #region 開啟與儲存檔案處理

            //自動調整欄位內容
            classCell.Worksheets[0].AutoFitColumns();

            //把Excel的各欄位加上特定格式
            for (int y = 4; y <= classCell.Worksheets[0].Cells.MaxDataColumn; y++)
            {
                Style style = classCell.Worksheets[0].Cells[1, y].GetStyle();
                style.SetBorder(BorderType.TopBorder,CellBorderType.Double,Color.Black);
                style.SetBorder(BorderType.BottomBorder, CellBorderType.Double, Color.Black);
                classCell.Worksheets[0].Cells[1, y].SetStyle(style);
            }

            try
            {
                SaveFileDialog1.Filter = "Excel (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";
                SaveFileDialog1.FileName = "班級出席狀況統計報表";

                Hide();
                if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    classCell.Save(SaveFileDialog1.FileName);
                    Process.Start(SaveFileDialog1.FileName);
                    Close();
                }
                else
                {
                    MessageBox.Show("檔案未儲存");
                }
            }
            catch
            {
                MessageBox.Show("檔案未儲存或檔案已被開啟");
            }

            #endregion

            Close();
        }

        private void checkBoxX1_CheckedChanged(object sender, EventArgs e)
        {
            foreach (ListViewItem each in listView1.Items)
            {
                each.Checked = checkBoxX1.Checked;
            }
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}