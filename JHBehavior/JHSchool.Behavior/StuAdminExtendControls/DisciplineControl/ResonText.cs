using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;

namespace JHSchool.Behavior.StuAdminExtendControls
{
    public partial class ResonText : BaseForm
    {
        public string Reson = "";


        public ResonText()
        {
            InitializeComponent();
        }

        public ResonText(string title)
        {
            InitializeComponent();

            this.Text = title;
        }

        private void ResonText_Load(object sender, EventArgs e)
        {

        }

        //Save
        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (textBoxX1.Text == "")
            {
                this.DialogResult = System.Windows.Forms.DialogResult.No;
            }
            else
            {
                Reson = textBoxX1.Text;
                this.DialogResult = System.Windows.Forms.DialogResult.Yes;
            }
        }
    }
}
