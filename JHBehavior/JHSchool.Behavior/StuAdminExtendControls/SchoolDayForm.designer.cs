namespace JHSchool.Behavior.StuAdminExtendControls
{
    partial class SchoolDayForm
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
            this.components = new System.ComponentModel.Container();
            this.dtBeginDate = new DevComponents.Editors.DateTimeAdv.DateTimeInput();
            this.dtEndDate = new DevComponents.Editors.DateTimeAdv.DateTimeInput();
            this.schoolStartDate = new DevComponents.DotNetBar.LabelX();
            this.schoolEndDate = new DevComponents.DotNetBar.LabelX();
            this.schoolyear = new DevComponents.DotNetBar.LabelX();
            this.txtG3Count = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
            this.txtG2Count = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.txtG1Count = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.SetButton = new DevComponents.DotNetBar.ButtonX();
            this.exitButton = new DevComponents.DotNetBar.ButtonX();
            this.advTree1 = new DevComponents.AdvTree.AdvTree();
            this.columnHeader1 = new DevComponents.AdvTree.ColumnHeader();
            this.columnHeader5 = new DevComponents.AdvTree.ColumnHeader();
            this.nodeConnector1 = new DevComponents.AdvTree.NodeConnector();
            this.elementStyle1 = new DevComponents.DotNetBar.ElementStyle();
            this.columnHeader3 = new DevComponents.AdvTree.ColumnHeader();
            this.columnHeader2 = new DevComponents.AdvTree.ColumnHeader();
            this.node2 = new DevComponents.AdvTree.Node();
            this.cell1 = new DevComponents.AdvTree.Cell();
            this.errorProvider1 = new System.Windows.Forms.ErrorProvider(this.components);
            this.groupPanel2 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.node1 = new DevComponents.AdvTree.Node();
            this.groupPanel3 = new DevComponents.DotNetBar.Controls.GroupPanel();
            ((System.ComponentModel.ISupportInitialize)(this.dtBeginDate)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtEndDate)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.advTree1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).BeginInit();
            this.groupPanel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // dtBeginDate
            // 
            this.dtBeginDate.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.dtBeginDate.BackgroundStyle.Class = "DateTimeInputBackground";
            this.dtBeginDate.ButtonDropDown.Shortcut = DevComponents.DotNetBar.eShortcut.AltDown;
            this.dtBeginDate.ButtonDropDown.Visible = true;
            this.dtBeginDate.ButtonFreeText.Checked = true;
            this.dtBeginDate.FreeTextEntryMode = true;
            this.dtBeginDate.Location = new System.Drawing.Point(106, 44);
            // 
            // 
            // 
            this.dtBeginDate.MonthCalendar.AnnuallyMarkedDates = new System.DateTime[0];
            // 
            // 
            // 
            this.dtBeginDate.MonthCalendar.BackgroundStyle.BackColor = System.Drawing.SystemColors.Window;
            this.dtBeginDate.MonthCalendar.ClearButtonVisible = true;
            // 
            // 
            // 
            this.dtBeginDate.MonthCalendar.CommandsBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground2;
            this.dtBeginDate.MonthCalendar.CommandsBackgroundStyle.BackColorGradientAngle = 90;
            this.dtBeginDate.MonthCalendar.CommandsBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground;
            this.dtBeginDate.MonthCalendar.CommandsBackgroundStyle.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.dtBeginDate.MonthCalendar.CommandsBackgroundStyle.BorderTopColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarDockedBorder;
            this.dtBeginDate.MonthCalendar.CommandsBackgroundStyle.BorderTopWidth = 1;
            this.dtBeginDate.MonthCalendar.DayNames = new string[] {
        "日",
        "一",
        "二",
        "三",
        "四",
        "五",
        "六"};
            this.dtBeginDate.MonthCalendar.DisplayMonth = new System.DateTime(2009, 5, 1, 0, 0, 0, 0);
            this.dtBeginDate.MonthCalendar.MarkedDates = new System.DateTime[0];
            this.dtBeginDate.MonthCalendar.MonthlyMarkedDates = new System.DateTime[0];
            // 
            // 
            // 
            this.dtBeginDate.MonthCalendar.NavigationBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.dtBeginDate.MonthCalendar.NavigationBackgroundStyle.BackColorGradientAngle = 90;
            this.dtBeginDate.MonthCalendar.NavigationBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.dtBeginDate.MonthCalendar.TodayButtonVisible = true;
            this.dtBeginDate.MonthCalendar.WeeklyMarkedDays = new System.DayOfWeek[0];
            this.dtBeginDate.Name = "dtBeginDate";
            this.dtBeginDate.Size = new System.Drawing.Size(125, 25);
            this.dtBeginDate.TabIndex = 1;
            this.dtBeginDate.ValueChanged += new System.EventHandler(this.dtBeginDate_ValueChanged);
            // 
            // dtEndDate
            // 
            this.dtEndDate.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.dtEndDate.BackgroundStyle.Class = "DateTimeInputBackground";
            this.dtEndDate.ButtonDropDown.Shortcut = DevComponents.DotNetBar.eShortcut.AltDown;
            this.dtEndDate.ButtonDropDown.Visible = true;
            this.dtEndDate.ButtonFreeText.Checked = true;
            this.dtEndDate.FreeTextEntryMode = true;
            this.dtEndDate.Location = new System.Drawing.Point(340, 44);
            // 
            // 
            // 
            this.dtEndDate.MonthCalendar.AnnuallyMarkedDates = new System.DateTime[0];
            // 
            // 
            // 
            this.dtEndDate.MonthCalendar.BackgroundStyle.BackColor = System.Drawing.SystemColors.Window;
            this.dtEndDate.MonthCalendar.ClearButtonVisible = true;
            // 
            // 
            // 
            this.dtEndDate.MonthCalendar.CommandsBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground2;
            this.dtEndDate.MonthCalendar.CommandsBackgroundStyle.BackColorGradientAngle = 90;
            this.dtEndDate.MonthCalendar.CommandsBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground;
            this.dtEndDate.MonthCalendar.CommandsBackgroundStyle.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.dtEndDate.MonthCalendar.CommandsBackgroundStyle.BorderTopColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarDockedBorder;
            this.dtEndDate.MonthCalendar.CommandsBackgroundStyle.BorderTopWidth = 1;
            this.dtEndDate.MonthCalendar.DayNames = new string[] {
        "日",
        "一",
        "二",
        "三",
        "四",
        "五",
        "六"};
            this.dtEndDate.MonthCalendar.DisplayMonth = new System.DateTime(2009, 5, 1, 0, 0, 0, 0);
            this.dtEndDate.MonthCalendar.MarkedDates = new System.DateTime[0];
            this.dtEndDate.MonthCalendar.MonthlyMarkedDates = new System.DateTime[0];
            // 
            // 
            // 
            this.dtEndDate.MonthCalendar.NavigationBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.dtEndDate.MonthCalendar.NavigationBackgroundStyle.BackColorGradientAngle = 90;
            this.dtEndDate.MonthCalendar.NavigationBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.dtEndDate.MonthCalendar.TodayButtonVisible = true;
            this.dtEndDate.MonthCalendar.WeeklyMarkedDays = new System.DayOfWeek[0];
            this.dtEndDate.Name = "dtEndDate";
            this.dtEndDate.Size = new System.Drawing.Size(125, 25);
            this.dtEndDate.TabIndex = 2;
            this.dtEndDate.ValueChanged += new System.EventHandler(this.dtEndDate_ValueChanged);
            // 
            // schoolStartDate
            // 
            this.schoolStartDate.BackColor = System.Drawing.Color.Transparent;
            this.schoolStartDate.Location = new System.Drawing.Point(12, 45);
            this.schoolStartDate.Name = "schoolStartDate";
            this.schoolStartDate.Size = new System.Drawing.Size(90, 23);
            this.schoolStartDate.TabIndex = 3;
            this.schoolStartDate.Text = "學期開始日期";
            // 
            // schoolEndDate
            // 
            this.schoolEndDate.BackColor = System.Drawing.Color.Transparent;
            this.schoolEndDate.Location = new System.Drawing.Point(249, 45);
            this.schoolEndDate.Name = "schoolEndDate";
            this.schoolEndDate.Size = new System.Drawing.Size(85, 22);
            this.schoolEndDate.TabIndex = 4;
            this.schoolEndDate.Text = "學期結束日期";
            // 
            // schoolyear
            // 
            this.schoolyear.BackColor = System.Drawing.Color.Transparent;
            this.schoolyear.Location = new System.Drawing.Point(12, 12);
            this.schoolyear.Name = "schoolyear";
            this.schoolyear.Size = new System.Drawing.Size(114, 26);
            this.schoolyear.TabIndex = 5;
            this.schoolyear.Text = "oo學年度oo學期";
            // 
            // txtG3Count
            // 
            // 
            // 
            // 
            this.txtG3Count.Border.Class = "TextBoxBorder";
            this.txtG3Count.Location = new System.Drawing.Point(356, 16);
            this.txtG3Count.Name = "txtG3Count";
            this.txtG3Count.Size = new System.Drawing.Size(44, 25);
            this.txtG3Count.TabIndex = 5;
            this.txtG3Count.Text = "0";
            // 
            // labelX3
            // 
            this.labelX3.AutoSize = true;
            this.labelX3.BackColor = System.Drawing.Color.Transparent;
            this.labelX3.Location = new System.Drawing.Point(285, 18);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(47, 21);
            this.labelX3.TabIndex = 4;
            this.labelX3.Text = "三年級";
            // 
            // txtG2Count
            // 
            // 
            // 
            // 
            this.txtG2Count.Border.Class = "TextBoxBorder";
            this.txtG2Count.Location = new System.Drawing.Point(204, 16);
            this.txtG2Count.Name = "txtG2Count";
            this.txtG2Count.Size = new System.Drawing.Size(44, 25);
            this.txtG2Count.TabIndex = 3;
            this.txtG2Count.Text = "0";
            // 
            // txtG1Count
            // 
            // 
            // 
            // 
            this.txtG1Count.Border.Class = "TextBoxBorder";
            this.txtG1Count.Location = new System.Drawing.Point(69, 16);
            this.txtG1Count.Name = "txtG1Count";
            this.txtG1Count.Size = new System.Drawing.Size(44, 25);
            this.txtG1Count.TabIndex = 2;
            this.txtG1Count.Text = "0";
            // 
            // labelX1
            // 
            this.labelX1.AutoSize = true;
            this.labelX1.Location = new System.Drawing.Point(146, 18);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(47, 21);
            this.labelX1.TabIndex = 1;
            this.labelX1.Text = "二年級";
            // 
            // labelX2
            // 
            this.labelX2.AutoSize = true;
            this.labelX2.Location = new System.Drawing.Point(11, 18);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(47, 21);
            this.labelX2.TabIndex = 0;
            this.labelX2.Text = "一年級";
            // 
            // SetButton
            // 
            this.SetButton.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.SetButton.BackColor = System.Drawing.Color.Transparent;
            this.SetButton.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.SetButton.Location = new System.Drawing.Point(286, 408);
            this.SetButton.Name = "SetButton";
            this.SetButton.Size = new System.Drawing.Size(75, 23);
            this.SetButton.TabIndex = 7;
            this.SetButton.Text = "儲存設定";
            this.SetButton.Click += new System.EventHandler(this.SetButton_Click);
            // 
            // exitButton
            // 
            this.exitButton.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.exitButton.BackColor = System.Drawing.Color.Transparent;
            this.exitButton.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.exitButton.Location = new System.Drawing.Point(383, 408);
            this.exitButton.Name = "exitButton";
            this.exitButton.Size = new System.Drawing.Size(75, 23);
            this.exitButton.TabIndex = 8;
            this.exitButton.Text = "離開";
            this.exitButton.Click += new System.EventHandler(this.btn_Exit);
            // 
            // advTree1
            // 
            this.advTree1.AccessibleRole = System.Windows.Forms.AccessibleRole.Outline;
            this.advTree1.AllowDrop = true;
            this.advTree1.BackColor = System.Drawing.SystemColors.Window;
            // 
            // 
            // 
            this.advTree1.BackgroundStyle.Class = "TreeBorderKey";
            this.advTree1.CellEdit = true;
            this.advTree1.Columns.Add(this.columnHeader1);
            this.advTree1.Columns.Add(this.columnHeader5);
            this.advTree1.LicenseKey = "F962CEC7-CD8F-4911-A9E9-CAB39962FC1F";
            this.advTree1.Location = new System.Drawing.Point(12, 83);
            this.advTree1.Name = "advTree1";
            this.advTree1.NodesConnector = this.nodeConnector1;
            this.advTree1.NodeStyle = this.elementStyle1;
            this.advTree1.PathSeparator = ";";
            this.advTree1.Size = new System.Drawing.Size(453, 227);
            this.advTree1.TabIndex = 0;
            this.advTree1.Text = "advTree1";
            // 
            // columnHeader1
            // 
            this.columnHeader1.Name = "columnHeader1";
            this.columnHeader1.Text = "上課日期";
            this.columnHeader1.Width.Absolute = 250;
            // 
            // columnHeader5
            // 
            this.columnHeader5.Name = "columnHeader5";
            this.columnHeader5.Text = "放假";
            this.columnHeader5.Width.Absolute = 100;
            // 
            // nodeConnector1
            // 
            this.nodeConnector1.LineColor = System.Drawing.SystemColors.ControlText;
            // 
            // elementStyle1
            // 
            this.elementStyle1.Name = "elementStyle1";
            this.elementStyle1.TextColor = System.Drawing.SystemColors.ControlText;
            // 
            // columnHeader3
            // 
            this.columnHeader3.Name = "columnHeader3";
            this.columnHeader3.Text = "放假";
            this.columnHeader3.Width.Absolute = 150;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Name = "columnHeader2";
            this.columnHeader2.Text = "上課日期";
            this.columnHeader2.Width.Absolute = 150;
            // 
            // node2
            // 
            this.node2.Cells.Add(this.cell1);
            this.node2.Expanded = true;
            this.node2.Name = "node2";
            this.node2.Text = "node2";
            // 
            // cell1
            // 
            this.cell1.CheckBoxStyle = DevComponents.DotNetBar.eCheckBoxStyle.RadioButton;
            this.cell1.CheckBoxVisible = true;
            this.cell1.Checked = true;
            this.cell1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cell1.Name = "cell1";
            this.cell1.StyleMouseOver = null;
            // 
            // errorProvider1
            // 
            this.errorProvider1.ContainerControl = this;
            // 
            // groupPanel2
            // 
            this.groupPanel2.BackColor = System.Drawing.Color.Transparent;
            this.groupPanel2.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel2.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel2.Location = new System.Drawing.Point(310, 237);
            this.groupPanel2.Name = "groupPanel2";
            this.groupPanel2.Size = new System.Drawing.Size(200, 100);
            // 
            // 
            // 
            this.groupPanel2.Style.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.groupPanel2.Style.BackColorGradientAngle = 90;
            this.groupPanel2.Style.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.groupPanel2.Style.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel2.Style.BorderBottomWidth = 1;
            this.groupPanel2.Style.BorderColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.groupPanel2.Style.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel2.Style.BorderLeftWidth = 1;
            this.groupPanel2.Style.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel2.Style.BorderRightWidth = 1;
            this.groupPanel2.Style.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel2.Style.BorderTopWidth = 1;
            this.groupPanel2.Style.CornerDiameter = 4;
            this.groupPanel2.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded;
            this.groupPanel2.Style.TextAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Center;
            this.groupPanel2.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.groupPanel2.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near;
            this.groupPanel2.TabIndex = 9;
            this.groupPanel2.Text = "groupPanel2";
            // 
            // node1
            // 
            this.node1.Expanded = true;
            this.node1.HostedControl = this.groupPanel2;
            this.node1.Name = "node1";
            this.node1.Text = "groupPanel2";
            // 
            // groupPanel3
            // 
            this.groupPanel3.BackColor = System.Drawing.Color.Transparent;
            this.groupPanel3.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel3.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel3.Controls.Add(this.txtG3Count);
            this.groupPanel3.Controls.Add(this.labelX2);
            this.groupPanel3.Controls.Add(this.labelX3);
            this.groupPanel3.Controls.Add(this.txtG1Count);
            this.groupPanel3.Controls.Add(this.txtG2Count);
            this.groupPanel3.Controls.Add(this.labelX1);
            this.groupPanel3.Location = new System.Drawing.Point(12, 328);
            this.groupPanel3.Name = "groupPanel3";
            this.groupPanel3.Size = new System.Drawing.Size(453, 74);
            // 
            // 
            // 
            this.groupPanel3.Style.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.groupPanel3.Style.BackColorGradientAngle = 90;
            this.groupPanel3.Style.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.groupPanel3.Style.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel3.Style.BorderBottomWidth = 1;
            this.groupPanel3.Style.BorderColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.groupPanel3.Style.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel3.Style.BorderLeftWidth = 1;
            this.groupPanel3.Style.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel3.Style.BorderRightWidth = 1;
            this.groupPanel3.Style.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel3.Style.BorderTopWidth = 1;
            this.groupPanel3.Style.CornerDiameter = 4;
            this.groupPanel3.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded;
            this.groupPanel3.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.groupPanel3.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near;
            this.groupPanel3.TabIndex = 9;
            this.groupPanel3.Text = "上課天數";
            // 
            // SchoolDayForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(477, 445);
            this.Controls.Add(this.groupPanel3);
            this.Controls.Add(this.exitButton);
            this.Controls.Add(this.SetButton);
            this.Controls.Add(this.schoolyear);
            this.Controls.Add(this.schoolEndDate);
            this.Controls.Add(this.schoolStartDate);
            this.Controls.Add(this.dtEndDate);
            this.Controls.Add(this.dtBeginDate);
            this.Controls.Add(this.advTree1);
            this.Name = "SchoolDayForm";
            this.Text = "上課天數設定";
            this.Load += new System.EventHandler(this.PeriodInSchool_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dtBeginDate)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtEndDate)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.advTree1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).EndInit();
            this.groupPanel3.ResumeLayout(false);
            this.groupPanel3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.Editors.DateTimeAdv.DateTimeInput dtBeginDate;
        private DevComponents.Editors.DateTimeAdv.DateTimeInput dtEndDate;
        private DevComponents.DotNetBar.LabelX schoolStartDate;
        private DevComponents.DotNetBar.LabelX schoolEndDate;
        private DevComponents.DotNetBar.LabelX schoolyear;
        private DevComponents.DotNetBar.ButtonX SetButton;
        private DevComponents.DotNetBar.ButtonX exitButton;
        private DevComponents.AdvTree.AdvTree advTree1;
        private DevComponents.AdvTree.ColumnHeader columnHeader1;
        private DevComponents.AdvTree.NodeConnector nodeConnector1;
        private DevComponents.DotNetBar.ElementStyle elementStyle1;
        private DevComponents.AdvTree.ColumnHeader columnHeader3;
        private DevComponents.AdvTree.ColumnHeader columnHeader2;
        private DevComponents.AdvTree.ColumnHeader columnHeader5;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.Controls.TextBoxX txtG2Count;
        private DevComponents.DotNetBar.Controls.TextBoxX txtG1Count;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.LabelX labelX3;
        private DevComponents.DotNetBar.Controls.TextBoxX txtG3Count;
        private DevComponents.AdvTree.Node node2;
        private DevComponents.AdvTree.Cell cell1;
        private System.Windows.Forms.ErrorProvider errorProvider1;
        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel3;
        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel2;
        private DevComponents.AdvTree.Node node1;
    }
}