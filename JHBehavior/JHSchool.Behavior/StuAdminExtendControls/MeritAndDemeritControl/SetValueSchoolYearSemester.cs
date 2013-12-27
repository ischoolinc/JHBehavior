using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;

namespace JHSchool.Behavior.StuAdminExtendControls.MeritAndDemeritControl
{
    public partial class SetValueSchoolYearSemester : BaseForm
    {
        public int _schoolYear { get; set; }
        public int _semester { get; set; }

        public SetValueSchoolYearSemester()
        {
            InitializeComponent();

            _schoolYear = integerInput1.Value = int.Parse(School.DefaultSchoolYear);
            _semester = integerInput2.Value = int.Parse(School.DefaultSemester);
        }

        public SetValueSchoolYearSemester(int schoolYear, int semester)
        {
            InitializeComponent();

            _schoolYear = integerInput1.Value = schoolYear;
            _semester = integerInput2.Value = semester;
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            _schoolYear = integerInput1.Value;
            _semester = integerInput2.Value;

            DialogResult dr = MsgBox.Show("按下確認,會立即儲存資料\n確認修改資料?", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2);
            if (dr == System.Windows.Forms.DialogResult.Yes)
            {
                this.DialogResult = System.Windows.Forms.DialogResult.Yes;
            }
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.No;
        }
    }
}
