using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using JHSchool.Data;
using FISCA.LogAgent;

namespace JHSchool.Behavior.StuAdminExtendControls.MeritAndDemeritControl
{
    public partial class ChangeTextForm : BaseForm
    {
        public string ChangeText = "";
        List<JHDisciplineRecord> _helper;

        public ChangeTextForm()
        {
            InitializeComponent();
        }

        public ChangeTextForm(List<JHDisciplineRecord> helper)
        {
            InitializeComponent();

            _helper = helper;
            textBoxX1.Text = _helper[0].Reason;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            DialogResult dr = MsgBox.Show("確認儲存事由修改?", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2);

            if (dr == System.Windows.Forms.DialogResult.Yes)
            {
                ChangeText = textBoxX1.Text.Trim();

                foreach (JHDisciplineRecord each in _helper)
                {
                    each.Reason = ChangeText;
                }

                try
                {
                    JHDiscipline.Update(_helper);

                    MsgBox.Show("資料儲存完成");
                    ApplicationLog.Log("獎懲批次修改", "修改", "批次修改事由資料\n" + "共" + _helper.Count + "筆資料");
                    this.DialogResult = System.Windows.Forms.DialogResult.Yes;
                }
                catch (Exception ex)
                {
                    MsgBox.Show("儲存發生錯誤:\n" + ex.Message);
                }
            }
        }
    }
}
