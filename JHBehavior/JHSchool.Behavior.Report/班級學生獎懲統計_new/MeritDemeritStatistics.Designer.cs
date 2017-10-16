namespace JHSchool.Behavior.Report.班級學生獎懲統計
{
    partial class MeritDemeritStatistics
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.listView1 = new System.Windows.Forms.ListView();
            this.SelectAllChecked = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.lbStartTxt = new DevComponents.DotNetBar.LabelX();
            this.btnPrint = new DevComponents.DotNetBar.ButtonX();
            this.lbEndTxt = new DevComponents.DotNetBar.LabelX();
            this.lbSetupTxt = new DevComponents.DotNetBar.LabelX();
            this.dtiStartDate = new DevComponents.Editors.DateTimeAdv.DateTimeInput();
            this.dtiEndDate = new DevComponents.Editors.DateTimeAdv.DateTimeInput();
            this.checkBoxX1 = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.rbRegisterDate = new System.Windows.Forms.RadioButton();
            this.rbStartDate = new System.Windows.Forms.RadioButton();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            ((System.ComponentModel.ISupportInitialize)(this.dtiStartDate)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtiEndDate)).BeginInit();
            this.SuspendLayout();
            // 
            // listView1
            // 
            this.listView1.Alignment = System.Windows.Forms.ListViewAlignment.Left;
            this.listView1.CheckBoxes = true;
            this.listView1.Location = new System.Drawing.Point(14, 9);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(109, 163);
            this.listView1.TabIndex = 6;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.SmallIcon;
            // 
            // SelectAllChecked
            // 
            this.SelectAllChecked.AutoSize = true;
            this.SelectAllChecked.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.SelectAllChecked.BackgroundStyle.Class = "";
            this.SelectAllChecked.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.SelectAllChecked.Checked = true;
            this.SelectAllChecked.CheckState = System.Windows.Forms.CheckState.Checked;
            this.SelectAllChecked.CheckValue = "Y";
            this.SelectAllChecked.Location = new System.Drawing.Point(14, 178);
            this.SelectAllChecked.Name = "SelectAllChecked";
            this.SelectAllChecked.Size = new System.Drawing.Size(80, 21);
            this.SelectAllChecked.TabIndex = 7;
            this.SelectAllChecked.Text = "選擇全部";
            this.SelectAllChecked.CheckedChanged += new System.EventHandler(this.SelectAllChecked_CheckedChanged);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(247, 197);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(60, 23);
            this.btnExit.TabIndex = 18;
            this.btnExit.Text = "取消";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // lbStartTxt
            // 
            this.lbStartTxt.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbStartTxt.BackgroundStyle.Class = "";
            this.lbStartTxt.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbStartTxt.Location = new System.Drawing.Point(129, 37);
            this.lbStartTxt.Name = "lbStartTxt";
            this.lbStartTxt.Size = new System.Drawing.Size(46, 23);
            this.lbStartTxt.TabIndex = 17;
            this.lbStartTxt.Text = "起始於";
            // 
            // btnPrint
            // 
            this.btnPrint.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPrint.BackColor = System.Drawing.Color.Transparent;
            this.btnPrint.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPrint.Location = new System.Drawing.Point(181, 197);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(60, 23);
            this.btnPrint.TabIndex = 16;
            this.btnPrint.Text = "列印";
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // lbEndTxt
            // 
            this.lbEndTxt.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbEndTxt.BackgroundStyle.Class = "";
            this.lbEndTxt.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbEndTxt.Location = new System.Drawing.Point(129, 68);
            this.lbEndTxt.Name = "lbEndTxt";
            this.lbEndTxt.Size = new System.Drawing.Size(46, 23);
            this.lbEndTxt.TabIndex = 14;
            this.lbEndTxt.Text = "結束於";
            // 
            // lbSetupTxt
            // 
            this.lbSetupTxt.AutoSize = true;
            this.lbSetupTxt.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbSetupTxt.BackgroundStyle.Class = "";
            this.lbSetupTxt.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbSetupTxt.Location = new System.Drawing.Point(129, 9);
            this.lbSetupTxt.Name = "lbSetupTxt";
            this.lbSetupTxt.Size = new System.Drawing.Size(127, 21);
            this.lbSetupTxt.TabIndex = 19;
            this.lbSetupTxt.Text = "設定獎懲日期區間：";
            // 
            // dtiStartDate
            // 
            this.dtiStartDate.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.dtiStartDate.BackgroundStyle.Class = "DateTimeInputBackground";
            this.dtiStartDate.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtiStartDate.ButtonDropDown.Shortcut = DevComponents.DotNetBar.eShortcut.AltDown;
            this.dtiStartDate.ButtonDropDown.Visible = true;
            this.dtiStartDate.IsPopupCalendarOpen = false;
            this.dtiStartDate.Location = new System.Drawing.Point(180, 36);
            // 
            // 
            // 
            this.dtiStartDate.MonthCalendar.AnnuallyMarkedDates = new System.DateTime[0];
            // 
            // 
            // 
            this.dtiStartDate.MonthCalendar.BackgroundStyle.BackColor = System.Drawing.SystemColors.Window;
            this.dtiStartDate.MonthCalendar.BackgroundStyle.Class = "";
            this.dtiStartDate.MonthCalendar.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtiStartDate.MonthCalendar.ClearButtonVisible = true;
            // 
            // 
            // 
            this.dtiStartDate.MonthCalendar.CommandsBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground2;
            this.dtiStartDate.MonthCalendar.CommandsBackgroundStyle.BackColorGradientAngle = 90;
            this.dtiStartDate.MonthCalendar.CommandsBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground;
            this.dtiStartDate.MonthCalendar.CommandsBackgroundStyle.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.dtiStartDate.MonthCalendar.CommandsBackgroundStyle.BorderTopColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarDockedBorder;
            this.dtiStartDate.MonthCalendar.CommandsBackgroundStyle.BorderTopWidth = 1;
            this.dtiStartDate.MonthCalendar.CommandsBackgroundStyle.Class = "";
            this.dtiStartDate.MonthCalendar.CommandsBackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtiStartDate.MonthCalendar.DayNames = new string[] {
        "日",
        "一",
        "二",
        "三",
        "四",
        "五",
        "六"};
            this.dtiStartDate.MonthCalendar.DisplayMonth = new System.DateTime(2010, 2, 1, 0, 0, 0, 0);
            this.dtiStartDate.MonthCalendar.MarkedDates = new System.DateTime[0];
            this.dtiStartDate.MonthCalendar.MonthlyMarkedDates = new System.DateTime[0];
            // 
            // 
            // 
            this.dtiStartDate.MonthCalendar.NavigationBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.dtiStartDate.MonthCalendar.NavigationBackgroundStyle.BackColorGradientAngle = 90;
            this.dtiStartDate.MonthCalendar.NavigationBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.dtiStartDate.MonthCalendar.NavigationBackgroundStyle.Class = "";
            this.dtiStartDate.MonthCalendar.NavigationBackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtiStartDate.MonthCalendar.TodayButtonVisible = true;
            this.dtiStartDate.MonthCalendar.WeeklyMarkedDays = new System.DayOfWeek[0];
            this.dtiStartDate.Name = "dtiStartDate";
            this.dtiStartDate.Size = new System.Drawing.Size(127, 25);
            this.dtiStartDate.TabIndex = 20;
            // 
            // dtiEndDate
            // 
            this.dtiEndDate.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.dtiEndDate.BackgroundStyle.Class = "DateTimeInputBackground";
            this.dtiEndDate.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtiEndDate.ButtonDropDown.Shortcut = DevComponents.DotNetBar.eShortcut.AltDown;
            this.dtiEndDate.ButtonDropDown.Visible = true;
            this.dtiEndDate.IsPopupCalendarOpen = false;
            this.dtiEndDate.Location = new System.Drawing.Point(180, 67);
            // 
            // 
            // 
            this.dtiEndDate.MonthCalendar.AnnuallyMarkedDates = new System.DateTime[0];
            // 
            // 
            // 
            this.dtiEndDate.MonthCalendar.BackgroundStyle.BackColor = System.Drawing.SystemColors.Window;
            this.dtiEndDate.MonthCalendar.BackgroundStyle.Class = "";
            this.dtiEndDate.MonthCalendar.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtiEndDate.MonthCalendar.ClearButtonVisible = true;
            // 
            // 
            // 
            this.dtiEndDate.MonthCalendar.CommandsBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground2;
            this.dtiEndDate.MonthCalendar.CommandsBackgroundStyle.BackColorGradientAngle = 90;
            this.dtiEndDate.MonthCalendar.CommandsBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground;
            this.dtiEndDate.MonthCalendar.CommandsBackgroundStyle.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.dtiEndDate.MonthCalendar.CommandsBackgroundStyle.BorderTopColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarDockedBorder;
            this.dtiEndDate.MonthCalendar.CommandsBackgroundStyle.BorderTopWidth = 1;
            this.dtiEndDate.MonthCalendar.CommandsBackgroundStyle.Class = "";
            this.dtiEndDate.MonthCalendar.CommandsBackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtiEndDate.MonthCalendar.DayNames = new string[] {
        "日",
        "一",
        "二",
        "三",
        "四",
        "五",
        "六"};
            this.dtiEndDate.MonthCalendar.DisplayMonth = new System.DateTime(2010, 2, 1, 0, 0, 0, 0);
            this.dtiEndDate.MonthCalendar.MarkedDates = new System.DateTime[0];
            this.dtiEndDate.MonthCalendar.MonthlyMarkedDates = new System.DateTime[0];
            // 
            // 
            // 
            this.dtiEndDate.MonthCalendar.NavigationBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.dtiEndDate.MonthCalendar.NavigationBackgroundStyle.BackColorGradientAngle = 90;
            this.dtiEndDate.MonthCalendar.NavigationBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.dtiEndDate.MonthCalendar.NavigationBackgroundStyle.Class = "";
            this.dtiEndDate.MonthCalendar.NavigationBackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtiEndDate.MonthCalendar.TodayButtonVisible = true;
            this.dtiEndDate.MonthCalendar.WeeklyMarkedDays = new System.DayOfWeek[0];
            this.dtiEndDate.Name = "dtiEndDate";
            this.dtiEndDate.Size = new System.Drawing.Size(127, 25);
            this.dtiEndDate.TabIndex = 21;
            // 
            // checkBoxX1
            // 
            this.checkBoxX1.AutoSize = true;
            this.checkBoxX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.checkBoxX1.BackgroundStyle.Class = "";
            this.checkBoxX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.checkBoxX1.Location = new System.Drawing.Point(14, 200);
            this.checkBoxX1.Name = "checkBoxX1";
            this.checkBoxX1.Size = new System.Drawing.Size(107, 21);
            this.checkBoxX1.TabIndex = 23;
            this.checkBoxX1.Text = "包含銷過記錄";
            // 
            // rbRegisterDate
            // 
            this.rbRegisterDate.BackColor = System.Drawing.Color.Transparent;
            this.rbRegisterDate.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.rbRegisterDate.Location = new System.Drawing.Point(180, 125);
            this.rbRegisterDate.Name = "rbRegisterDate";
            this.rbRegisterDate.Size = new System.Drawing.Size(107, 21);
            this.rbRegisterDate.TabIndex = 0;
            this.rbRegisterDate.Text = "獎懲登錄日期";
            this.rbRegisterDate.UseVisualStyleBackColor = false;
            // 
            // rbStartDate
            // 
            this.rbStartDate.AutoSize = true;
            this.rbStartDate.BackColor = System.Drawing.Color.Transparent;
            this.rbStartDate.Checked = true;
            this.rbStartDate.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.rbStartDate.Location = new System.Drawing.Point(180, 98);
            this.rbStartDate.Name = "rbStartDate";
            this.rbStartDate.Size = new System.Drawing.Size(104, 21);
            this.rbStartDate.TabIndex = 24;
            this.rbStartDate.TabStop = true;
            this.rbStartDate.Text = "獎懲發生日期";
            this.rbStartDate.UseVisualStyleBackColor = false;
            // 
            // labelX1
            // 
            this.labelX1.AutoSize = true;
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(129, 152);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(167, 21);
            this.labelX1.TabIndex = 25;
            this.labelX1.Text = "說明：本功能僅統計一般生";
            // 
            // MeritDemeritStatistics
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(322, 232);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.rbStartDate);
            this.Controls.Add(this.rbRegisterDate);
            this.Controls.Add(this.checkBoxX1);
            this.Controls.Add(this.dtiEndDate);
            this.Controls.Add(this.dtiStartDate);
            this.Controls.Add(this.lbSetupTxt);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.lbStartTxt);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.lbEndTxt);
            this.Controls.Add(this.SelectAllChecked);
            this.Controls.Add(this.listView1);
            this.DoubleBuffered = true;
            this.Name = "MeritDemeritStatistics";
            this.Text = "班級獎懲狀況統計";
            this.Load += new System.EventHandler(this.MeritDemeritStatistics_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dtiStartDate)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtiEndDate)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView listView1;
        private DevComponents.DotNetBar.Controls.CheckBoxX SelectAllChecked;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.LabelX lbStartTxt;
        private DevComponents.DotNetBar.ButtonX btnPrint;
        private DevComponents.DotNetBar.LabelX lbEndTxt;
        private DevComponents.DotNetBar.LabelX lbSetupTxt;
        private DevComponents.Editors.DateTimeAdv.DateTimeInput dtiStartDate;
        private DevComponents.Editors.DateTimeAdv.DateTimeInput dtiEndDate;
        private DevComponents.DotNetBar.Controls.CheckBoxX checkBoxX1;
        private System.Windows.Forms.RadioButton rbRegisterDate;
        private System.Windows.Forms.RadioButton rbStartDate;
        private DevComponents.DotNetBar.LabelX labelX1;
    }
}