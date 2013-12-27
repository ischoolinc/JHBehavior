using SmartSchool.Common;
namespace JHSchool.Behavior.ClassExtendControls.Ribbon
{
    partial class DemeritStatistic
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
            this.groupPanel1 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
            this.panelEx2 = new DevComponents.DotNetBar.PanelEx();
            this.labelX8 = new DevComponents.DotNetBar.LabelX();
            this.labelX9 = new DevComponents.DotNetBar.LabelX();
            this.txtC = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.labelX10 = new DevComponents.DotNetBar.LabelX();
            this.txtB = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.txtA = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.cboSemester = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.cboSchoolYear = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.rbType2 = new System.Windows.Forms.RadioButton();
            this.rbType1 = new System.Windows.Forms.RadioButton();
            this.error = new System.Windows.Forms.ErrorProvider(this.components);
            this.groupPanel1.SuspendLayout();
            this.panelEx2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.error)).BeginInit();
            this.SuspendLayout();
            // 
            // groupPanel1
            // 
            this.groupPanel1.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel1.Controls.Add(this.checkBox1);
            this.groupPanel1.Controls.Add(this.labelX3);
            this.groupPanel1.Controls.Add(this.panelEx2);
            this.groupPanel1.Controls.Add(this.cboSemester);
            this.groupPanel1.Controls.Add(this.cboSchoolYear);
            this.groupPanel1.Controls.Add(this.labelX2);
            this.groupPanel1.Controls.Add(this.labelX1);
            this.groupPanel1.Controls.Add(this.rbType2);
            this.groupPanel1.Controls.Add(this.rbType1);
            this.groupPanel1.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.groupPanel1.Location = new System.Drawing.Point(0, 0);
            this.groupPanel1.Name = "groupPanel1";
            this.groupPanel1.Size = new System.Drawing.Size(391, 393);
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
            this.groupPanel1.TabIndex = 1;
            this.groupPanel1.Text = "懲戒表現累計";
            // 
            // checkBox1
            // 
            this.checkBox1.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.checkBox1.AutoSize = true;
            this.checkBox1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(66)))), ((int)(((byte)(133)))));
            this.checkBox1.Location = new System.Drawing.Point(40, 179);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(105, 21);
            this.checkBox1.TabIndex = 9;
            this.checkBox1.Text = "是否功過相抵";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.Visible = false;
            // 
            // labelX3
            // 
            this.labelX3.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.labelX3.BackColor = System.Drawing.Color.Transparent;
            this.labelX3.Font = new System.Drawing.Font("微軟正黑體", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labelX3.Location = new System.Drawing.Point(59, 242);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(266, 79);
            this.labelX3.TabIndex = 8;
            this.labelX3.Text = "<font color=\"red\">說明</font><br/>\r\n懲戒表現之累計針對學生未銷過之懲戒紀錄進行\r\n統計，若學生之懲戒紀錄已銷過，則不列入累計\r\n數" +
                "字統計。";
            this.labelX3.TextLineAlignment = System.Drawing.StringAlignment.Near;
            // 
            // panelEx2
            // 
            this.panelEx2.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.panelEx2.CanvasColor = System.Drawing.SystemColors.Control;
            this.panelEx2.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.panelEx2.Controls.Add(this.labelX8);
            this.panelEx2.Controls.Add(this.labelX9);
            this.panelEx2.Controls.Add(this.txtC);
            this.panelEx2.Controls.Add(this.labelX10);
            this.panelEx2.Controls.Add(this.txtB);
            this.panelEx2.Controls.Add(this.txtA);
            this.panelEx2.Location = new System.Drawing.Point(33, 101);
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
            this.panelEx2.TabIndex = 7;
            // 
            // labelX8
            // 
            this.labelX8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.labelX8.BackColor = System.Drawing.Color.Transparent;
            this.labelX8.Location = new System.Drawing.Point(115, 18);
            this.labelX8.Name = "labelX8";
            this.labelX8.Size = new System.Drawing.Size(34, 23);
            this.labelX8.TabIndex = 1;
            this.labelX8.Text = "小過";
            // 
            // labelX9
            // 
            this.labelX9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.labelX9.BackColor = System.Drawing.Color.Transparent;
            this.labelX9.Location = new System.Drawing.Point(20, 18);
            this.labelX9.Name = "labelX9";
            this.labelX9.Size = new System.Drawing.Size(34, 23);
            this.labelX9.TabIndex = 1;
            this.labelX9.Text = "大過";
            // 
            // txtC
            // 
            this.txtC.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            // 
            // 
            // 
            this.txtC.Border.Class = "TextBoxBorder";
            this.txtC.Location = new System.Drawing.Point(256, 17);
            this.txtC.Name = "txtC";
            this.txtC.Size = new System.Drawing.Size(37, 25);
            this.txtC.TabIndex = 4;
            // 
            // labelX10
            // 
            this.labelX10.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.labelX10.BackColor = System.Drawing.Color.Transparent;
            this.labelX10.Location = new System.Drawing.Point(210, 18);
            this.labelX10.Name = "labelX10";
            this.labelX10.Size = new System.Drawing.Size(34, 23);
            this.labelX10.TabIndex = 1;
            this.labelX10.Text = "警告";
            // 
            // txtB
            // 
            this.txtB.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            // 
            // 
            // 
            this.txtB.Border.Class = "TextBoxBorder";
            this.txtB.Location = new System.Drawing.Point(161, 17);
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
            this.txtA.Location = new System.Drawing.Point(66, 17);
            this.txtA.Name = "txtA";
            this.txtA.Size = new System.Drawing.Size(37, 25);
            this.txtA.TabIndex = 4;
            // 
            // cboSemester
            // 
            this.cboSemester.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.cboSemester.DisplayMember = "Text";
            this.cboSemester.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboSemester.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboSemester.FormattingEnabled = true;
            this.cboSemester.ItemHeight = 19;
            this.cboSemester.Location = new System.Drawing.Point(308, 26);
            this.cboSemester.Name = "cboSemester";
            this.cboSemester.Size = new System.Drawing.Size(49, 25);
            this.cboSemester.TabIndex = 2;
            // 
            // cboSchoolYear
            // 
            this.cboSchoolYear.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.cboSchoolYear.DisplayMember = "Text";
            this.cboSchoolYear.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboSchoolYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboSchoolYear.FormattingEnabled = true;
            this.cboSchoolYear.ItemHeight = 19;
            this.cboSchoolYear.Location = new System.Drawing.Point(221, 26);
            this.cboSchoolYear.Name = "cboSchoolYear";
            this.cboSchoolYear.Size = new System.Drawing.Size(49, 25);
            this.cboSchoolYear.TabIndex = 2;
            // 
            // labelX2
            // 
            this.labelX2.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.labelX2.BackColor = System.Drawing.Color.Transparent;
            this.labelX2.Location = new System.Drawing.Point(276, 27);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(36, 23);
            this.labelX2.TabIndex = 1;
            this.labelX2.Text = "學期";
            // 
            // labelX1
            // 
            this.labelX1.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            this.labelX1.Location = new System.Drawing.Point(170, 27);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(50, 23);
            this.labelX1.TabIndex = 1;
            this.labelX1.Text = "學年度";
            // 
            // rbType2
            // 
            this.rbType2.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.rbType2.AutoSize = true;
            this.rbType2.BackColor = System.Drawing.Color.Transparent;
            this.rbType2.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.rbType2.Location = new System.Drawing.Point(33, 65);
            this.rbType2.Name = "rbType2";
            this.rbType2.Size = new System.Drawing.Size(143, 21);
            this.rbType2.TabIndex = 0;
            this.rbType2.TabStop = true;
            this.rbType2.Text = "各學期懲戒表現累計";
            this.rbType2.UseVisualStyleBackColor = false;
            // 
            // rbType1
            // 
            this.rbType1.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.rbType1.AutoSize = true;
            this.rbType1.BackColor = System.Drawing.Color.Transparent;
            this.rbType1.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.rbType1.Location = new System.Drawing.Point(33, 28);
            this.rbType1.Name = "rbType1";
            this.rbType1.Size = new System.Drawing.Size(130, 21);
            this.rbType1.TabIndex = 0;
            this.rbType1.TabStop = true;
            this.rbType1.Text = "單一學期懲戒表現";
            this.rbType1.UseVisualStyleBackColor = false;
            // 
            // error
            // 
            this.error.ContainerControl = this;
            // 
            // DemeritStatistic
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.groupPanel1);
            this.Name = "DemeritStatistic";
            this.Size = new System.Drawing.Size(391, 393);
            this.groupPanel1.ResumeLayout(false);
            this.groupPanel1.PerformLayout();
            this.panelEx2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.error)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel1;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboSemester;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboSchoolYear;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.LabelX labelX1;
        private System.Windows.Forms.RadioButton rbType2;
        private System.Windows.Forms.RadioButton rbType1;
        private System.Windows.Forms.ErrorProvider error;
        private DevComponents.DotNetBar.PanelEx panelEx2;
        private DevComponents.DotNetBar.LabelX labelX8;
        private DevComponents.DotNetBar.LabelX labelX9;
        private DevComponents.DotNetBar.Controls.TextBoxX txtC;
        private DevComponents.DotNetBar.LabelX labelX10;
        private DevComponents.DotNetBar.Controls.TextBoxX txtB;
        private DevComponents.DotNetBar.Controls.TextBoxX txtA;
        private DevComponents.DotNetBar.LabelX labelX3;
        private System.Windows.Forms.CheckBox checkBox1;
    }
}
