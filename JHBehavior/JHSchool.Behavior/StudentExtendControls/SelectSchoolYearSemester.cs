using FISCA.Presentation.Controls;
using Framework.Security;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace JHSchool.Behavior.StudentExtendControls
{
    public partial class SelectSchoolYearSemester : BaseForm
    {
        string _StudentID;
        int _SchoolYear, _Semester;
        internal static FeatureAce _UserPermission;

        public SelectSchoolYearSemester(string studentID, FeatureAce UserPermission)
        {
            InitializeComponent();
            _StudentID = studentID;
            _UserPermission = UserPermission;
        }

        private void SelectSchoolYearSemester_Load(object sender, EventArgs e)
        {
            int schoolyear = int.Parse(K12.Data.School.DefaultSchoolYear);
            int semester = int.Parse(K12.Data.School.DefaultSemester);

            for (int i = schoolyear - 3; i < schoolyear + 3; i++)
            {
                comboBoxEx1.Items.Add(i.ToString());
            }

            comboBoxEx2.Items.Add("1");
            comboBoxEx2.Items.Add("2");

            comboBoxEx1.Text = schoolyear.ToString();
            comboBoxEx2.Text = semester.ToString();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            _SchoolYear = int.Parse(comboBoxEx1.Text);
            _Semester = int.Parse(comboBoxEx2.Text);

            AttendanceUnifyForm AttendanceForm = new AttendanceUnifyForm(_StudentID, _SchoolYear, _Semester, _UserPermission);
            AttendanceForm.ShowDialog();
            this.Close();
        }
    }
}
