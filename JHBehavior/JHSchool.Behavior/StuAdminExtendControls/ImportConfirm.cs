using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Framework;
//using SmartSchool.Common;

namespace JHSchool.Behavior.StuAdminExtendControls
{
    public partial class ImportConfirm : FISCA.Presentation.Controls.BaseForm
    {
        private bool _overwrite = false;
        public bool Overwrite
        {
            get { return _overwrite; }
        }

        public ImportConfirm()
        {
            InitializeComponent();
        }

        private void buttonX3_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            _overwrite = true;
            this.DialogResult = DialogResult.OK;
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            if (FISCA.Presentation.Controls.MsgBox.Show("如果原有的德行評語代碼表與匯入的德行評語代碼表中有重複的評語代碼，\n則會以後來匯入的德行評語代碼表為主。是否繼續匯入動作？", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                _overwrite = false;
                this.DialogResult = DialogResult.OK;
            }
            else
                this.DialogResult = DialogResult.Cancel;
        }
    }
}