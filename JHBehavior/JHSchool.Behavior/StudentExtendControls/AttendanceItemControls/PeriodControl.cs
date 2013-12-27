using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar.Controls;
using DevComponents.DotNetBar;

namespace JHSchool.Behavior.StudentExtendControls
{
    public partial class PeriodControl : UserControl
    {
        public PeriodControl()
        {
            InitializeComponent();
            this.Font = Framework.DotNetBar.FontStyles.General;
            this.Width = 45;
        }

        public LabelX Label
        {
            get { return label; }
        }

        public TextBoxX TextBox
        {
            get { return textBox; }
        }
    }
}
