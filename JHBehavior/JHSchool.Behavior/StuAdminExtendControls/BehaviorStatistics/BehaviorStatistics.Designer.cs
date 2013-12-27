namespace JHSchool.Behavior.StuAdminExtendControls.BehaviorStatistics
{
    partial class BehaviorStatistics
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
            this.btnSTART = new DevComponents.DotNetBar.ButtonX();
            this.btnClose = new DevComponents.DotNetBar.ButtonX();
            this.cbSchoolYear = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.cbSemester = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.lbSchoolYear = new DevComponents.DotNetBar.LabelX();
            this.lbSemester = new DevComponents.DotNetBar.LabelX();
            this.lbGradeYear = new DevComponents.DotNetBar.LabelX();
            this.cbGradeYear = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.errorProvider1 = new System.Windows.Forms.ErrorProvider(this.components);
            this.errorProvider2 = new System.Windows.Forms.ErrorProvider(this.components);
            this.errorProvider3 = new System.Windows.Forms.ErrorProvider(this.components);
            this.groupPanel1 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.groupPanel2 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider3)).BeginInit();
            this.groupPanel1.SuspendLayout();
            this.groupPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnSTART
            // 
            this.btnSTART.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSTART.BackColor = System.Drawing.Color.Transparent;
            this.btnSTART.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSTART.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.btnSTART.Location = new System.Drawing.Point(132, 331);
            this.btnSTART.Name = "btnSTART";
            this.btnSTART.Size = new System.Drawing.Size(75, 23);
            this.btnSTART.TabIndex = 0;
            this.btnSTART.Text = "開始計算";
            this.btnSTART.Click += new System.EventHandler(this.btnSTART_Click);
            // 
            // btnClose
            // 
            this.btnClose.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnClose.BackColor = System.Drawing.Color.Transparent;
            this.btnClose.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnClose.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.btnClose.Location = new System.Drawing.Point(229, 331);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 1;
            this.btnClose.Text = "離開";
            this.btnClose.Click += new System.EventHandler(this.BTNColse_Click);
            // 
            // cbSchoolYear
            // 
            this.cbSchoolYear.DisplayMember = "Text";
            this.cbSchoolYear.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbSchoolYear.FormattingEnabled = true;
            this.cbSchoolYear.ItemHeight = 19;
            this.cbSchoolYear.Location = new System.Drawing.Point(91, 23);
            this.cbSchoolYear.Name = "cbSchoolYear";
            this.cbSchoolYear.Size = new System.Drawing.Size(64, 25);
            this.cbSchoolYear.TabIndex = 2;
            this.cbSchoolYear.TextChanged += new System.EventHandler(this.cbSchoolYear_TextChanged);
            // 
            // cbSemester
            // 
            this.cbSemester.DisplayMember = "Text";
            this.cbSemester.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbSemester.FormattingEnabled = true;
            this.cbSemester.ItemHeight = 19;
            this.cbSemester.Location = new System.Drawing.Point(197, 23);
            this.cbSemester.Name = "cbSemester";
            this.cbSemester.Size = new System.Drawing.Size(64, 25);
            this.cbSemester.TabIndex = 3;
            this.cbSemester.TextChanged += new System.EventHandler(this.cbSemester_TextChanged);
            // 
            // lbSchoolYear
            // 
            this.lbSchoolYear.AutoSize = true;
            this.lbSchoolYear.BackColor = System.Drawing.Color.Transparent;
            this.lbSchoolYear.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.lbSchoolYear.Location = new System.Drawing.Point(23, 25);
            this.lbSchoolYear.Name = "lbSchoolYear";
            this.lbSchoolYear.Size = new System.Drawing.Size(70, 21);
            this.lbSchoolYear.TabIndex = 4;
            this.lbSchoolYear.Text = "學 年 度 ：";
            // 
            // lbSemester
            // 
            this.lbSemester.AutoSize = true;
            this.lbSemester.BackColor = System.Drawing.Color.Transparent;
            this.lbSemester.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.lbSemester.Location = new System.Drawing.Point(159, 25);
            this.lbSemester.Name = "lbSemester";
            this.lbSemester.Size = new System.Drawing.Size(34, 21);
            this.lbSemester.TabIndex = 5;
            this.lbSemester.Text = "學期";
            // 
            // lbGradeYear
            // 
            this.lbGradeYear.AutoSize = true;
            this.lbGradeYear.BackColor = System.Drawing.Color.Transparent;
            this.lbGradeYear.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.lbGradeYear.Location = new System.Drawing.Point(22, 16);
            this.lbGradeYear.Name = "lbGradeYear";
            this.lbGradeYear.Size = new System.Drawing.Size(47, 21);
            this.lbGradeYear.TabIndex = 7;
            this.lbGradeYear.Text = "年級：";
            // 
            // cbGradeYear
            // 
            this.cbGradeYear.DisplayMember = "Text";
            this.cbGradeYear.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbGradeYear.FormattingEnabled = true;
            this.cbGradeYear.ItemHeight = 19;
            this.cbGradeYear.Location = new System.Drawing.Point(87, 14);
            this.cbGradeYear.Name = "cbGradeYear";
            this.cbGradeYear.Size = new System.Drawing.Size(64, 25);
            this.cbGradeYear.TabIndex = 8;
            this.cbGradeYear.TextChanged += new System.EventHandler(this.cbGradeYear_TextChanged);
            // 
            // errorProvider1
            // 
            this.errorProvider1.ContainerControl = this;
            // 
            // errorProvider2
            // 
            this.errorProvider2.ContainerControl = this;
            // 
            // errorProvider3
            // 
            this.errorProvider3.ContainerControl = this;
            // 
            // groupPanel1
            // 
            this.groupPanel1.BackColor = System.Drawing.Color.Transparent;
            this.groupPanel1.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel1.Controls.Add(this.labelX2);
            this.groupPanel1.Controls.Add(this.cbGradeYear);
            this.groupPanel1.Controls.Add(this.lbGradeYear);
            this.groupPanel1.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.groupPanel1.Location = new System.Drawing.Point(12, 12);
            this.groupPanel1.Name = "groupPanel1";
            this.groupPanel1.Size = new System.Drawing.Size(290, 112);
            // 
            // 
            // 
            this.groupPanel1.Style.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.groupPanel1.Style.BackColorGradientAngle = 90;
            this.groupPanel1.Style.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.groupPanel1.Style.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderBottomWidth = 1;
            this.groupPanel1.Style.BorderColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.groupPanel1.Style.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderLeftWidth = 1;
            this.groupPanel1.Style.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderRightWidth = 1;
            this.groupPanel1.Style.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderTopWidth = 1;
            this.groupPanel1.Style.CornerDiameter = 4;
            this.groupPanel1.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded;
            this.groupPanel1.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.groupPanel1.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near;
            this.groupPanel1.TabIndex = 12;
            this.groupPanel1.Text = "範圍";
            // 
            // labelX2
            // 
            this.labelX2.AutoSize = true;
            this.labelX2.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.labelX2.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.labelX2.Location = new System.Drawing.Point(18, 47);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(163, 21);
            this.labelX2.TabIndex = 9;
            this.labelX2.Text = "(必須為系統中存在之年級)";
            // 
            // groupPanel2
            // 
            this.groupPanel2.BackColor = System.Drawing.Color.Transparent;
            this.groupPanel2.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel2.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel2.Controls.Add(this.labelX1);
            this.groupPanel2.Controls.Add(this.cbSemester);
            this.groupPanel2.Controls.Add(this.cbSchoolYear);
            this.groupPanel2.Controls.Add(this.lbSchoolYear);
            this.groupPanel2.Controls.Add(this.lbSemester);
            this.groupPanel2.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.groupPanel2.Location = new System.Drawing.Point(12, 130);
            this.groupPanel2.Name = "groupPanel2";
            this.groupPanel2.Size = new System.Drawing.Size(290, 129);
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
            this.groupPanel2.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.groupPanel2.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near;
            this.groupPanel2.TabIndex = 13;
            this.groupPanel2.Text = "條件";
            // 
            // labelX1
            // 
            this.labelX1.AutoSize = true;
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            this.labelX1.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.labelX1.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.labelX1.Location = new System.Drawing.Point(18, 56);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(176, 21);
            this.labelX1.TabIndex = 14;
            this.labelX1.Text = "(請選擇要計算之學年度學期)";
            // 
            // timer1
            // 
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.BackColor = System.Drawing.Color.Transparent;
            this.radioButton2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.radioButton2.Location = new System.Drawing.Point(12, 293);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(224, 21);
            this.radioButton2.TabIndex = 21;
            this.radioButton2.Text = "無論有無明細內容,均進行覆蓋計算";
            this.radioButton2.UseVisualStyleBackColor = false;
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.BackColor = System.Drawing.Color.Transparent;
            this.radioButton1.Checked = true;
            this.radioButton1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.radioButton1.Location = new System.Drawing.Point(12, 268);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(268, 21);
            this.radioButton1.TabIndex = 20;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "無缺曠/獎/懲明細之學生,不列入計算(預設)";
            this.radioButton1.UseVisualStyleBackColor = false;
            // 
            // BehaviorStatistics
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(316, 366);
            this.Controls.Add(this.radioButton2);
            this.Controls.Add(this.radioButton1);
            this.Controls.Add(this.groupPanel2);
            this.Controls.Add(this.groupPanel1);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnSTART);
            this.Name = "BehaviorStatistics";
            this.Text = "缺曠獎懲學期結算";
            this.Load += new System.EventHandler(this.BehaviorStatistics_Load);
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider3)).EndInit();
            this.groupPanel1.ResumeLayout(false);
            this.groupPanel1.PerformLayout();
            this.groupPanel2.ResumeLayout(false);
            this.groupPanel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX btnSTART;
        private DevComponents.DotNetBar.ButtonX btnClose;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbSchoolYear;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbSemester;
        private DevComponents.DotNetBar.LabelX lbSchoolYear;
        private DevComponents.DotNetBar.LabelX lbSemester;
        private DevComponents.DotNetBar.LabelX lbGradeYear;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbGradeYear;
        private System.Windows.Forms.ErrorProvider errorProvider1;
        private System.Windows.Forms.ErrorProvider errorProvider2;
        private System.Windows.Forms.ErrorProvider errorProvider3;
        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel2;
        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel1;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.LabelX labelX1;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.RadioButton radioButton1;
    }
}