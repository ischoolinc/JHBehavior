using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
//using SmartSchool.Common;
using FISCA.DSAUtil;
using SmartSchool.Common;
using Framework;
using JHSchool.Feature.Legacy;
using JHSchool.Data;
using FISCA.LogAgent;
//using SmartSchool;

namespace JHSchool.Behavior.StuAdminExtendControls.MeritAndDemeritControl
{
    public partial class ModifyForm : FISCA.Presentation.Controls.BaseForm
    {
        private string _merit_flag;
        private List<JHDisciplineRecord> _helper;
        private StringBuilder sb = new StringBuilder();

        public string NewReason
        {
            get { return txtNewReason.Text; }
        }

        public List<JHDisciplineRecord> Helper
        {
            get { return _helper; }
        }

        public ModifyForm(List<JHDisciplineRecord> helper)
        {
            InitializeComponent();

            _helper = helper;

            _merit_flag = helper[0].MeritFlag;

            intSchoolYear.Value = helper[0].SchoolYear;
            intSemester.Value = helper[0].Semester;
            txtNewReason.Text = helper[0].Reason;

            sb.AppendLine("已進行批次修改獎懲資料");
            sb.AppendLine("依據第一筆資料狀態：");
            sb.AppendLine("學年度「" + helper[0].SchoolYear + "」學期「" + helper[0].Semester + "」");

            if (_merit_flag == "1")
            {
                sb.AppendLine("大功「" + helper[0].MeritA.Value + "」小功「" + helper[0].MeritB.Value + "」嘉獎「" + helper[0].MeritC.Value + "」");
                lblA.Text = "大功";
                lblB.Text = "小功";
                lblC.Text = "嘉獎";
                txtA.Text = "" + helper[0].MeritA;
                txtB.Text = "" + helper[0].MeritB;
                txtC.Text = "" + helper[0].MeritC;
            }
            else if (_merit_flag == "0")
            {
                sb.AppendLine("大過「" + helper[0].DemeritA.Value + "」小過「" + helper[0].DemeritB.Value + "」警告「" + helper[0].DemeritC.Value + "」");
                lblA.Text = "大過";
                lblB.Text = "小過";
                lblC.Text = "警告";
                txtA.Text = "" + helper[0].DemeritA;
                txtB.Text = "" + helper[0].DemeritB;
                txtC.Text = "" + helper[0].DemeritC;
            }
            else if (_merit_flag == "2")
            {
                lblA.Enabled = lblB.Enabled = lblC.Enabled = false;
                txtA.Enabled = txtB.Enabled = txtC.Enabled = false;
            }
            sb.AppendLine("事由「" + txtNewReason.Text);
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!ValidateDisciplineCount()) return;

            if (txtNewReason.Text.Trim() == "")
            {
                DialogResult dr = FISCA.Presentation.Controls.MsgBox.Show("事由未輸入,是否繼續進行儲存操作?", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2);
                if (dr == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }
            }

            DialogResult KJ = FISCA.Presentation.Controls.MsgBox.Show("是否修改獎懲內容?", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2);

            if (KJ == DialogResult.Yes)
            {
                sb.AppendLine("\n已批次修改為：");
                sb.AppendLine("學年度「" + intSchoolYear.Value + "」學期「" + intSemester.Value + "」");
                if (_merit_flag == "1")
                {
                    sb.AppendLine("大功「" + txtA.Text + "」小功「" + txtB.Text + "」嘉獎「" + txtC.Text + "」");
                }
                else if (_merit_flag == "0")
                {
                    sb.AppendLine("大過「" + txtA.Text + "」小過「" + txtB.Text + "」警告「" + txtC.Text + "」");
                }
                sb.AppendLine("事由「" + txtNewReason.Text + "\n");

                foreach (JHDisciplineRecord each in _helper)
                {
                    sb.AppendLine("學生「" + each.Student.Name + "」班級「" + (each.Student.Class != null ? each.Student.Class.Name : "") + "」座號「" + (each.Student.SeatNo.HasValue ? each.Student.SeatNo.Value.ToString() : "") + "」");
                    each.Reason = txtNewReason.Text;
                    each.SchoolYear = intSchoolYear.Value;
                    each.Semester = intSemester.Value;

                    if (_merit_flag == "1")
                    {
                        each.MeritA = int.Parse(txtA.Text);
                        each.MeritB = int.Parse(txtB.Text);
                        each.MeritC = int.Parse(txtC.Text);
                    }
                    else if (_merit_flag == "0")
                    {
                        each.DemeritA = int.Parse(txtA.Text);
                        each.DemeritB = int.Parse(txtB.Text);
                        each.DemeritC = int.Parse(txtC.Text);
                    }
                }

                try
                {
                    JHDiscipline.Update(_helper);
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("批次更改資料失敗!!。" + ex);
                    return;
                }

                ApplicationLog.Log("獎懲批次修改", "修改", sb.ToString() + "\n批次修改共「" + _helper.Count + "」筆。");
            }

            this.DialogResult = DialogResult.OK;
        }

        private bool ValidateDisciplineCount()
        {
            bool valid = true;
            errorProvider.Clear();
            int v;
            if (!int.TryParse(txtA.Text, out v))
            {
                errorProvider.SetIconAlignment(txtA, ErrorIconAlignment.MiddleLeft);
                errorProvider.SetError(txtA, "必須為數字");
                valid = false;
            }
            if (!int.TryParse(txtB.Text, out v))
            {
                errorProvider.SetIconAlignment(txtB, ErrorIconAlignment.MiddleLeft);
                errorProvider.SetError(txtB, "必須為數字");
                valid = false;
            }
            if (!int.TryParse(txtC.Text, out v))
            {
                errorProvider.SetIconAlignment(txtC, ErrorIconAlignment.MiddleLeft);
                errorProvider.SetError(txtC, "必須為數字");
                valid = false;
            }
            if (txtA.Text == "0" && txtB.Text == "0" && txtC.Text == "0")
            {
                FISCA.Presentation.Controls.MsgBox.Show("您未輸入任何資料!!");
                errorProvider.SetIconAlignment(txtA, ErrorIconAlignment.MiddleLeft);
                errorProvider.SetError(txtA, "您未輸入任何資料");
                errorProvider.SetIconAlignment(txtB, ErrorIconAlignment.MiddleLeft);
                errorProvider.SetError(txtB, "您未輸入任何資料");
                errorProvider.SetIconAlignment(txtC, ErrorIconAlignment.MiddleLeft);
                errorProvider.SetError(txtC, "您未輸入任何資料");
                valid = false;
            }
            return valid;
        }
    }
}