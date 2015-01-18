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
            intSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
            intSemester.Value = int.Parse(K12.Data.School.DefaultSemester);
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            _SchoolYear = intSchoolYear.Value;
            _Semester = intSemester.Value;

            AttendanceUnifyForm AttendanceForm = new AttendanceUnifyForm(_StudentID, _SchoolYear, _Semester, _UserPermission);
            AttendanceForm.ShowDialog();
            this.Close();
        }
    }
}
