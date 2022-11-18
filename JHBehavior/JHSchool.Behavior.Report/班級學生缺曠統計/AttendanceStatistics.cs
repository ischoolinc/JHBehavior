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

            #region �B�z���O

            cd = K12.Data.School.Configuration["�Z�ťX�ʮu���p�έp��"];
            textBoxX3.Text = cd["�C��`��"];

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

            #region �B�z�`��

            int gtRoden;
            if (!int.TryParse(textBoxX3.Text, out gtRoden))
            {
                MessageBox.Show("�`�����e�D�Ʀr���A!");
                this.Enabled = true;
                return;
            }

            cd["�C��`��"] = textBoxX3.Text;
            cd.Save();

            #endregion

            #region Excel�إ�

            Workbook classCell = new Workbook(new MemoryStream(ProjectResource.�Z�žǥͯ��m�έp),new LoadOptions(LoadFormat.Excel97To2003));

            //�x�s����
            SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

            List<ClassRecord> classList = K12.Data.Class.SelectByIDs(K12.Presentation.NLDPanels.Class.SelectedSource);

            classList = SortClassIndex.K12Data_ClassRecord(classList);

            #endregion

            #region �إߦr��&�d��

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

            #region �B�z���

            if (dateTimeInput1.Value > dateTimeInput2.Value)
            {
                MessageBox.Show("�}�l�ɶ��j�󵲧��ɶ�!!");
                return;
            }


            int dayNum = 0; //���o����Ѽ�
            DateTime Temp = dateTimeInput1.Value;
            while (Temp <= dateTimeInput2.Value)
            {
                if (Temp.DayOfWeek == DayOfWeek.Saturday || Temp.DayOfWeek == DayOfWeek.Sunday) //��/�餣�C�J�X�u�v�p��
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
                    if (student.Status == StudentRecord.StudentStatus.�@�� || student.Status == StudentRecord.StudentStatus.����)
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

                #region �B�z���m���e


                foreach (AttendanceRecord AttInfo in AttendanceList)
                {
                    foreach (AttendancePeriod AttPer in AttInfo.PeriodDetail)
                    {
                        //dateTimeInput2.Value.AddDays(1); //�h�@��??

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

                //�p��X�u���v����

                //Att                           �`���q��
                //classRecord.Students.Count    �H��
                //gtRoden                       �@�Ѫ��`��
                //dayNum                        �D����Ѽ�
                if (studentList.Count != 0)
                {
                    TaClass = 1 - ((double)Att / (studentList.Count * gtRoden * dayNum));
                }
                TaClass = TaClass * 100;
                classCell.Worksheets[0].Cells[nX, DicNum].PutValue(Math.Round(TaClass, 2, MidpointRounding.AwayFromZero) + "%");
                classCell.Worksheets[0].Cells[1, DicNum].PutValue("�X�u�v");


                TaClass = 0;
                Att = 0;
                nX++;

            }

            this.Enabled = true;

            #region �}�һP�x�s�ɮ׳B�z

            //�۰ʽվ���줺�e
            classCell.Worksheets[0].AutoFitColumns();

            //��Excel���U���[�W�S�w�榡
            for (int y = 4; y <= classCell.Worksheets[0].Cells.MaxDataColumn; y++)
            {
                Style style = classCell.Worksheets[0].Cells[1, y].GetStyle();
                style.SetBorder(BorderType.TopBorder,CellBorderType.Double,Color.Black);
                style.SetBorder(BorderType.BottomBorder, CellBorderType.Double, Color.Black);
                classCell.Worksheets[0].Cells[1, y].SetStyle(style);
            }

            try
            {
                SaveFileDialog1.Filter = "Excel (*.xlsx)|*.xlsx|�Ҧ��ɮ� (*.*)|*.*";
                SaveFileDialog1.FileName = "�Z�ťX�u���p�έp����";

                Hide();
                if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    classCell.Save(SaveFileDialog1.FileName);
                    Process.Start(SaveFileDialog1.FileName);
                    Close();
                }
                else
                {
                    MessageBox.Show("�ɮץ��x�s");
                }
            }
            catch
            {
                MessageBox.Show("�ɮץ��x�s���ɮפw�Q�}��");
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