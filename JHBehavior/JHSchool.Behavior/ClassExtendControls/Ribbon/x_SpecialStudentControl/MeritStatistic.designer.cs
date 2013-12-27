using SmartSchool.Common;
namespace JHSchool.Behavior.ClassExtendControls.Ribbon
{
    partial class MeritStatistic
    {
        /// <summary> 
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該公開 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 元件設計工具產生的程式碼

        /// <summary> 
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改這個方法的內容。
        ///
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupPanel1 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.cboSemester = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.labelX4 = new DevComponents.DotNetBar.LabelX();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
            this.cboSchoolYear = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.panelEx2 = new DevComponents.DotNetBar.PanelEx();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.labelX6 = new DevComponents.DotNetBar.LabelX();
            this.txtC = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.txtB = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.txtA = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.rbType2 = new System.Windows.Forms.RadioButton();
            this.rbType1 = new System.Windows.Forms.RadioButton();
            this.rbNoUnclearedDemert = new System.Windows.Forms.RadioButton();
            this.rbAll = new System.Windows.Forms.RadioButton();
            this.rbNoDemert = new System.Windows.Forms.RadioButton();
            this.error = new System.Windows.Forms.ErrorProvider(this.components);
            this.panel1.SuspendLayout();
            this.groupPanel1.SuspendLayout();
            this.panelEx2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.error)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.groupPanel1);
            this.panel1.Controls.Add(this.rbNoUnclearedDemert);
            this.panel1.Controls.Add(this.rbAll);
            this.panel1.Controls.Add(this.rbNoDemert);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(391, 393);
            this.panel1.TabIndex = 0;
            // 
            // groupPanel1
            // 
            this.groupPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.groupPanel1.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel1.Controls.Add(this.cboSemester);
            this.groupPanel1.Controls.Add(this.labelX4);
            this.groupPanel1.Controls.Add(this.labelX3);
            this.groupPanel1.Controls.Add(this.cboSchoolYear);
            this.groupPanel1.Controls.Add(this.panelEx2);
            this.groupPanel1.Controls.Add(this.rbType2);
            this.groupPanel1.Controls.Add(this.rbType1);
            this.groupPanel1.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.groupPanel1.Location = new System.Drawing.Point(0, 1);
            this.groupPanel1.Name = "groupPanel1";
            this.groupPanel1.Size = new System.Drawing.Size(391, 186);
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
            this.groupPanel1.TabIndex = 6;
            this.groupPanel1.Text = "獎勵表現累計";
            // 
            // cboSemester
            // 
            this.cboSemester.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.cboSemester.DisplayMember = "Text";
            this.cboSemester.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboSemester.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboSemester.FormattingEnabled = true;
            this.cboSemester.ItemHeight = 19;
            this.cboSemester.Location = new System.Drawing.Point(295, 11);
            this.cboSemester.Name = "cboSemester";
            this.cboSemester.Size = new System.Drawing.Size(39, 25);
            this.cboSemester.TabIndex = 7;
            // 
            // labelX4
            // 
            this.labelX4.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.labelX4.BackColor = System.Drawing.Color.Transparent;
            this.labelX4.Location = new System.Drawing.Point(253, 12);
            this.labelX4.Name = "labelX4";
            this.labelX4.Size = new System.Drawing.Size(36, 23);
            this.labelX4.TabIndex = 1;
            this.labelX4.Text = "學期";
            // 
            // labelX3
            // 
            this.labelX3.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.labelX3.BackColor = System.Drawing.Color.Transparent;
            this.labelX3.Location = new System.Drawing.Point(143, 12);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(47, 23);
            this.labelX3.TabIndex = 1;
            this.labelX3.Text = "學年度";
            // 
            // cboSchoolYear
            // 
            this.cboSchoolYear.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.cboSchoolYear.DisplayMember = "Text";
            this.cboSchoolYear.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboSchoolYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboSchoolYear.FormattingEnabled = true;
            this.cboSchoolYear.ItemHeight = 19;
            this.cboSchoolYear.Location = new System.Drawing.Point(196, 11);
            this.cboSchoolYear.Name = "cboSchoolYear";
            this.cboSchoolYear.Size = new System.Drawing.Size(51, 25);
            this.cboSchoolYear.TabIndex = 7;
            // 
            // panelEx2
            // 
            this.panelEx2.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.panelEx2.CanvasColor = System.Drawing.SystemColors.Control;
            this.panelEx2.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.panelEx2.Controls.Add(this.labelX2);
            this.panelEx2.Controls.Add(this.labelX6);
            this.panelEx2.Controls.Add(this.txtC);
            this.panelEx2.Controls.Add(this.labelX1);
            this.panelEx2.Controls.Add(this.txtB);
            this.panelEx2.Controls.Add(this.txtA);
            this.panelEx2.Location = new System.Drawing.Point(33, 68);
            this.panelEx2.Name = "panelEx2";
            this.panelEx2.Size = new System.Drawing.Size(312, 59);
            this.panelEx2.Style.Alignment = System.Drawing.StringAlignment.Center;
            this.panelEx2.Style.BackColor1.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.panelEx2.Style.BackColor2.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.panelEx2.Style.Border = DevComponents.DotNetBar.eBorderType.SingleLine;
            this.panelEx2.Style.BorderColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.panelEx2.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded;
            this.panelEx2.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.panelEx2.Style.GradientAngle = 90;
            this.panelEx2.TabIndex = 6;
            // 
            // labelX2
            // 
            this.labelX2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.labelX2.BackColor = System.Drawing.Color.Transparent;
            this.labelX2.Location = new System.Drawing.Point(114, 18);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(34, 23);
            this.labelX2.TabIndex = 1;
            this.labelX2.Text = "小功";
            // 
            // labelX6
            // 
            this.labelX6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.labelX6.BackColor = System.Drawing.Color.Transparent;
            this.labelX6.Location = new System.Drawing.Point(19, 18);
            this.labelX6.Name = "labelX6";
            this.labelX6.Size = new System.Drawing.Size(34, 23);
            this.labelX6.TabIndex = 1;
            this.labelX6.Text = "大功";
            // 
            // txtC
            // 
            this.txtC.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            // 
            // 
            // 
            this.txtC.Border.Class = "TextBoxBorder";
            this.txtC.Location = new System.Drawing.Point(255, 17);
            this.txtC.Name = "txtC";
            this.txtC.Size = new System.Drawing.Size(37, 25);
            this.txtC.TabIndex = 4;
            // 
            // labelX1
            // 
            this.labelX1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            this.labelX1.Location = new System.Drawing.Point(209, 18);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(34, 23);
            this.labelX1.TabIndex = 1;
            this.labelX1.Text = "嘉獎";
            // 
            // txtB
            // 
            this.txtB.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            // 
            // 
            // 
            this.txtB.Border.Class = "TextBoxBorder";
            this.txtB.Location = new System.Drawing.Point(160, 17);
            this.txtB.Name = "txtB";
            this.txtB.Size = new System.Drawing.Size(37, 25);
            this.txtB.TabIndex = 4;
            // 
            // txtA
            // 
            this.txtA.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            // 
            // 
            // 
            this.txtA.Border.Class = "TextBoxBorder";
            this.txtA.Location = new System.Drawing.Point(65, 17);
            this.txtA.Name = "txtA";
            this.txtA.Size = new System.Drawing.Size(37, 25);
            this.txtA.TabIndex = 4;
            // 
            // rbType2
            // 
            this.rbType2.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.rbType2.AutoSize = true;
            this.rbType2.BackColor = System.Drawing.Color.Transparent;
            this.rbType2.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.rbType2.Location = new System.Drawing.Point(33, 40);
            this.rbType2.Name = "rbType2";
            this.rbType2.Size = new System.Drawing.Size(104, 21);
            this.rbType2.TabIndex = 0;
            this.rbType2.TabStop = true;
            this.rbType2.Text = "至本學期合計";
            this.rbType2.UseVisualStyleBackColor = false;
            // 
            // rbType1
            // 
            this.rbType1.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.rbType1.AutoSize = true;
            this.rbType1.BackColor = System.Drawing.Color.Transparent;
            this.rbType1.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.rbType1.Location = new System.Drawing.Point(33, 13);
            this.rbType1.Name = "rbType1";
            this.rbType1.Size = new System.Drawing.Size(104, 21);
            this.rbType1.TabIndex = 0;
            this.rbType1.TabStop = true;
            this.rbType1.Text = "單一學期合計";
            this.rbType1.UseVisualStyleBackColor = false;
            // 
            // rbNoUnclearedDemert
            // 
            this.rbNoUnclearedDemert.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.rbNoUnclearedDemert.AutoSize = true;
            this.rbNoUnclearedDemert.BackColor = System.Drawing.Color.Transparent;
            this.rbNoUnclearedDemert.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.rbNoUnclearedDemert.Location = new System.Drawing.Point(36, 257);
            this.rbNoUnclearedDemert.Name = "rbNoUnclearedDemert";
            this.rbNoUnclearedDemert.Size = new System.Drawing.Size(273, 21);
            this.rbNoUnclearedDemert.TabIndex = 5;
            this.rbNoUnclearedDemert.Text = "若學生有未銷過之懲戒記錄，則不列入清單";
            this.rbNoUnclearedDemert.UseVisualStyleBackColor = false;
            // 
            // rbAll
            // 
            this.rbAll.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.rbAll.AutoSize = true;
            this.rbAll.BackColor = System.Drawing.Color.Transparent;
            this.rbAll.Checked = true;
            this.rbAll.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.rbAll.Location = new System.Drawing.Point(36, 203);
            this.rbAll.Name = "rbAll";
            this.rbAll.Size = new System.Drawing.Size(182, 21);
            this.rbAll.TabIndex = 3;
            this.rbAll.TabStop = true;
            this.rbAll.Text = "忽略學生懲戒記錄進行統計";
            this.rbAll.UseVisualStyleBackColor = false;
            // 
            // rbNoDemert
            // 
            this.rbNoDemert.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.rbNoDemert.AutoSize = true;
            this.rbNoDemert.BackColor = System.Drawing.Color.Transparent;
            this.rbNoDemert.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.rbNoDemert.Location = new System.Drawing.Point(36, 230);
            this.rbNoDemert.Name = "rbNoDemert";
            this.rbNoDemert.Size = new System.Drawing.Size(195, 21);
            this.rbNoDemert.TabIndex = 4;
            this.rbNoDemert.Text = "有懲戒記錄之學生不列入清單";
            this.rbNoDemert.UseVisualStyleBackColor = false;
            // 
            // error
            // 
            this.error.ContainerControl = this;
            // 
            // MeritStatistic
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Transparent;
            this.Controls.Add(this.panel1);
            this.Name = "MeritStatistic";
            this.Size = new System.Drawing.Size(391, 393);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.groupPanel1.ResumeLayout(false);
            this.groupPanel1.PerformLayout();
            this.panelEx2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.error)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ErrorProvider error;
        private System.Windows.Forms.Panel panel1;
        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel1;
        private DevComponents.DotNetBar.Controls.TextBoxX txtA;
        private DevComponents.DotNetBar.LabelX labelX6;
        private System.Windows.Forms.RadioButton rbType2;
        private System.Windows.Forms.RadioButton rbType1;
        private System.Windows.Forms.RadioButton rbNoUnclearedDemert;
        private System.Windows.Forms.RadioButton rbAll;
        private System.Windows.Forms.RadioButton rbNoDemert;
        private DevComponents.DotNetBar.Controls.TextBoxX txtC;
        private DevComponents.DotNetBar.Controls.TextBoxX txtB;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.PanelEx panelEx2;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboSemester;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboSchoolYear;
        private DevComponents.DotNetBar.LabelX labelX3;
        private DevComponents.DotNetBar.LabelX labelX4;
    }
}
