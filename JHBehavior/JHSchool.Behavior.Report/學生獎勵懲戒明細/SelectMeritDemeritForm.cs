using System;
using System.Windows.Forms;
using FISCA.Presentation.Controls;

namespace JHSchool.Behavior.Report
{
    public partial class SelectMeritDemeritForm : BaseForm
    {
        public SelectMeritDemeritForm()
        {
            InitializeComponent();

            comboBoxEx1.SelectedIndex = 0;

            #region �Ǧ~�׾Ǵ�
            string schoolYear = K12.Data.School.DefaultSchoolYear;
            cbSchoolYear.Text = schoolYear;
            cbSchoolYear.Items.Add((int.Parse(schoolYear) - 2).ToString());
            cbSchoolYear.Items.Add((int.Parse(schoolYear) - 1).ToString());
            cbSchoolYear.Items.Add((int.Parse(schoolYear)).ToString());

            string semester = K12.Data.School.DefaultSemester;
            cbSemester.Text = semester;
            cbSemester.Items.Add("1");
            cbSemester.Items.Add("2");
            #endregion

            string TimmeInput1 = DateTime.Now.AddMonths(-1).ToString("yyyy/MM/dd");
            dateTimeInput1.Text = TimmeInput1;

            string TimmeInput2 = DateTime.Now.ToString("yyyy/MM/dd");
            dateTimeInput2.Text = TimmeInput2;
        }

        #region �̤���٬O�̾Ǵ������A SelectDayOrSchoolYear

        public bool SelectDayOrSchoolYear
        {
            get
            {
                if (radioButton1.Checked)
                {
                    return true; //�̤��
                }
                else
                {
                    return false; //�̾Ǧ~�׾Ǵ�
                }
            }
        } 

        #endregion

        #region �̤����O�_�Ѧҵn������٬O�o�ͤ�� SetupTime

        public bool SetupTime
        {
            get
            {
                if (comboBoxEx1.SelectedIndex == 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }


        #endregion

        #region �Ǧ~�׾Ǵ����� SchoolYear,Semester

        public string SchoolYear
        {
            get
            {
                return cbSchoolYear.Text;
            }
        }

        public string Semester
        {
            get
            {
                return cbSemester.Text;
            }
        }

        /// <summary>
        /// �C�L�Ҧ��Ǧ~��
        /// </summary>
        public bool checkBoxX1Bool
        {
            get
            {
                return checkBoxX1.Checked;
            }
        }

        /// <summary>
        /// �O�_�ư��w�P�L���
        /// </summary>
        public bool checkBoxX2Bool
        {
            get
            {
                return checkBoxX2.Checked;
            }
        }

        /// <summary>
        /// �O�_�ư��w�P�L���
        /// </summary>
        public bool checkBoxX3Bool
        {
            get
            {
                return checkBoxX3.Checked;
            }
        }

        #endregion

        #region ������� StartDay,EndDay
        public DateTime StartDateTime //���o�}�l���
        {
            get
            {
                return dateTimeInput1.Value;
            }
        }

        public DateTime EndDateTime //���o�������
        {
            get
            {
                return dateTimeInput2.Value;
            }
        }
        #endregion

        protected virtual void buttonX1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void checkBoxX1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxX1.Checked)
            {
                cbSchoolYear.Enabled = false;
                cbSemester.Enabled = false;
            }
            else
            {
                cbSchoolYear.Enabled = true;
                cbSemester.Enabled = true;
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            comboBoxEx1.Enabled = radioButton1.Checked;
            dateTimeInput1.Enabled = radioButton1.Checked;
            dateTimeInput2.Enabled = radioButton1.Checked;
            labelX3.Enabled = radioButton1.Checked;
            labelX4.Enabled = radioButton1.Checked;
            labelX5.Enabled = radioButton1.Checked;
            checkBoxX1.Checked = false;
            cbSchoolYear.Enabled = false;
            cbSemester.Enabled = false;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            cbSchoolYear.Enabled = radioButton3.Checked;
            cbSemester.Enabled = radioButton3.Checked;
            labelX1.Enabled = radioButton3.Checked;
            labelX2.Enabled = radioButton3.Checked;
            checkBoxX1.Enabled = radioButton3.Checked;
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}