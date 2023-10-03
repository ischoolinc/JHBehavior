namespace JHSchool.Behavior.StudentExtendControls
{
    partial class AttendanceUnifyForm
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.lbHelp1 = new DevComponents.DotNetBar.LabelX();
            this.groupPanel4 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.lbHelp2 = new DevComponents.DotNetBar.LabelX();
            this.dgvAttendance = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.btnSaveAttendanceStatistics = new DevComponents.DotNetBar.ButtonX();
            this.groupPanel3 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.listViewAttendance = new JHSchool.Behavior.Legacy.ListViewEx();
            this.btnAttendanceNew = new DevComponents.DotNetBar.ButtonX();
            this.btnAttendanceDelete = new DevComponents.DotNetBar.ButtonX();
            this.buttonX4 = new DevComponents.DotNetBar.ButtonX();
            this.btnAttendanceEdit = new DevComponents.DotNetBar.ButtonX();
            this.tabItem2 = new DevComponents.DotNetBar.TabItem(this.components);
            this.btnExitAll = new DevComponents.DotNetBar.ButtonX();
            this.groupPanel4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAttendance)).BeginInit();
            this.groupPanel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // lbHelp1
            // 
            this.lbHelp1.AutoSize = true;
            this.lbHelp1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbHelp1.BackgroundStyle.Class = "";
            this.lbHelp1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbHelp1.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbHelp1.Location = new System.Drawing.Point(7, 5);
            this.lbHelp1.Name = "lbHelp1";
            this.lbHelp1.Size = new System.Drawing.Size(82, 26);
            this.lbHelp1.TabIndex = 29;
            this.lbHelp1.Text = "基本說明..";
            // 
            // groupPanel4
            // 
            this.groupPanel4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupPanel4.BackColor = System.Drawing.Color.Transparent;
            this.groupPanel4.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel4.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel4.Controls.Add(this.lbHelp2);
            this.groupPanel4.Controls.Add(this.dgvAttendance);
            this.groupPanel4.Controls.Add(this.btnSaveAttendanceStatistics);
            this.groupPanel4.Location = new System.Drawing.Point(12, 261);
            this.groupPanel4.Name = "groupPanel4";
            this.groupPanel4.Size = new System.Drawing.Size(569, 238);
            // 
            // 
            // 
            this.groupPanel4.Style.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.groupPanel4.Style.BackColorGradientAngle = 90;
            this.groupPanel4.Style.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.groupPanel4.Style.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel4.Style.BorderBottomWidth = 1;
            this.groupPanel4.Style.BorderColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.groupPanel4.Style.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel4.Style.BorderLeftWidth = 1;
            this.groupPanel4.Style.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel4.Style.BorderRightWidth = 1;
            this.groupPanel4.Style.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel4.Style.BorderTopWidth = 1;
            this.groupPanel4.Style.Class = "";
            this.groupPanel4.Style.CornerDiameter = 4;
            this.groupPanel4.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded;
            this.groupPanel4.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.groupPanel4.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near;
            // 
            // 
            // 
            this.groupPanel4.StyleMouseDown.Class = "";
            this.groupPanel4.StyleMouseDown.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            // 
            // 
            // 
            this.groupPanel4.StyleMouseOver.Class = "";
            this.groupPanel4.StyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.groupPanel4.TabIndex = 26;
            this.groupPanel4.Text = "缺曠統計";
            // 
            // lbHelp2
            // 
            this.lbHelp2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lbHelp2.AutoSize = true;
            // 
            // 
            // 
            this.lbHelp2.BackgroundStyle.Class = "";
            this.lbHelp2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbHelp2.Location = new System.Drawing.Point(180, 183);
            this.lbHelp2.Name = "lbHelp2";
            this.lbHelp2.Size = new System.Drawing.Size(181, 21);
            this.lbHelp2.TabIndex = 23;
            this.lbHelp2.Text = "說明：白色欄位為可調整內容";
            // 
            // dgvAttendance
            // 
            this.dgvAttendance.AllowUserToAddRows = false;
            this.dgvAttendance.AllowUserToDeleteRows = false;
            this.dgvAttendance.AllowUserToResizeRows = false;
            this.dgvAttendance.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvAttendance.BackgroundColor = System.Drawing.Color.White;
            this.dgvAttendance.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvAttendance.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgvAttendance.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgvAttendance.Location = new System.Drawing.Point(10, 8);
            this.dgvAttendance.Name = "dgvAttendance";
            this.dgvAttendance.RowHeadersVisible = false;
            this.dgvAttendance.RowTemplate.Height = 24;
            this.dgvAttendance.Size = new System.Drawing.Size(543, 167);
            this.dgvAttendance.TabIndex = 7;
            this.dgvAttendance.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvAttendance_CellEndEdit);
            this.dgvAttendance.CurrentCellDirtyStateChanged += new System.EventHandler(this.dgvAttendance_CurrentCellDirtyStateChanged);
            // 
            // btnSaveAttendanceStatistics
            // 
            this.btnSaveAttendanceStatistics.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSaveAttendanceStatistics.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnSaveAttendanceStatistics.AutoSize = true;
            this.btnSaveAttendanceStatistics.BackColor = System.Drawing.Color.Transparent;
            this.btnSaveAttendanceStatistics.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSaveAttendanceStatistics.Location = new System.Drawing.Point(11, 181);
            this.btnSaveAttendanceStatistics.Name = "btnSaveAttendanceStatistics";
            this.btnSaveAttendanceStatistics.Size = new System.Drawing.Size(158, 25);
            this.btnSaveAttendanceStatistics.TabIndex = 22;
            this.btnSaveAttendanceStatistics.Text = "儲存缺曠手動調整統計值";
            this.btnSaveAttendanceStatistics.Click += new System.EventHandler(this.btnSaveAttendanceStatistics_Click);
            // 
            // groupPanel3
            // 
            this.groupPanel3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupPanel3.BackColor = System.Drawing.Color.Transparent;
            this.groupPanel3.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel3.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel3.Controls.Add(this.listViewAttendance);
            this.groupPanel3.Controls.Add(this.btnAttendanceNew);
            this.groupPanel3.Controls.Add(this.btnAttendanceDelete);
            this.groupPanel3.Controls.Add(this.buttonX4);
            this.groupPanel3.Controls.Add(this.btnAttendanceEdit);
            this.groupPanel3.Location = new System.Drawing.Point(12, 37);
            this.groupPanel3.Name = "groupPanel3";
            this.groupPanel3.Size = new System.Drawing.Size(569, 218);
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
            this.groupPanel3.Style.Class = "";
            this.groupPanel3.Style.CornerDiameter = 4;
            this.groupPanel3.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded;
            this.groupPanel3.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.groupPanel3.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near;
            // 
            // 
            // 
            this.groupPanel3.StyleMouseDown.Class = "";
            this.groupPanel3.StyleMouseDown.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            // 
            // 
            // 
            this.groupPanel3.StyleMouseOver.Class = "";
            this.groupPanel3.StyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.groupPanel3.TabIndex = 26;
            this.groupPanel3.Text = "缺曠明細";
            // 
            // listViewAttendance
            // 
            this.listViewAttendance.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            // 
            // 
            // 
            this.listViewAttendance.Border.Class = "ListViewBorder";
            this.listViewAttendance.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.listViewAttendance.FullRowSelect = true;
            this.listViewAttendance.HideSelection = false;
            this.listViewAttendance.Location = new System.Drawing.Point(10, 6);
            this.listViewAttendance.Name = "listViewAttendance";
            this.listViewAttendance.Size = new System.Drawing.Size(543, 149);
            this.listViewAttendance.TabIndex = 11;
            this.listViewAttendance.UseCompatibleStateImageBehavior = false;
            this.listViewAttendance.View = System.Windows.Forms.View.Details;
            this.listViewAttendance.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.listViewAttendance_MouseDoubleClick);
            // 
            // btnAttendanceNew
            // 
            this.btnAttendanceNew.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnAttendanceNew.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnAttendanceNew.BackColor = System.Drawing.Color.Transparent;
            this.btnAttendanceNew.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnAttendanceNew.Location = new System.Drawing.Point(11, 161);
            this.btnAttendanceNew.Name = "btnAttendanceNew";
            this.btnAttendanceNew.Size = new System.Drawing.Size(75, 23);
            this.btnAttendanceNew.TabIndex = 2;
            this.btnAttendanceNew.Text = "新增缺曠";
            this.btnAttendanceNew.Click += new System.EventHandler(this.btnAttendanceNew_Click);
            // 
            // btnAttendanceDelete
            // 
            this.btnAttendanceDelete.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnAttendanceDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnAttendanceDelete.BackColor = System.Drawing.Color.Transparent;
            this.btnAttendanceDelete.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnAttendanceDelete.Location = new System.Drawing.Point(173, 161);
            this.btnAttendanceDelete.Name = "btnAttendanceDelete";
            this.btnAttendanceDelete.Size = new System.Drawing.Size(75, 23);
            this.btnAttendanceDelete.TabIndex = 4;
            this.btnAttendanceDelete.Text = "刪除缺曠";
            this.btnAttendanceDelete.Click += new System.EventHandler(this.btnAttendanceDelete_Click);
            // 
            // buttonX4
            // 
            this.buttonX4.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.buttonX4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonX4.BackColor = System.Drawing.Color.Transparent;
            this.buttonX4.Location = new System.Drawing.Point(11, 161);
            this.buttonX4.Name = "buttonX4";
            this.buttonX4.Size = new System.Drawing.Size(75, 23);
            this.buttonX4.TabIndex = 5;
            this.buttonX4.Text = "檢視";
            // 
            // btnAttendanceEdit
            // 
            this.btnAttendanceEdit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnAttendanceEdit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnAttendanceEdit.BackColor = System.Drawing.Color.Transparent;
            this.btnAttendanceEdit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnAttendanceEdit.Location = new System.Drawing.Point(92, 161);
            this.btnAttendanceEdit.Name = "btnAttendanceEdit";
            this.btnAttendanceEdit.Size = new System.Drawing.Size(75, 23);
            this.btnAttendanceEdit.TabIndex = 3;
            this.btnAttendanceEdit.Text = "修改缺曠";
            this.btnAttendanceEdit.Click += new System.EventHandler(this.btnAttendanceEdit_Click);
            // 
            // tabItem2
            // 
            this.tabItem2.Name = "tabItem2";
            this.tabItem2.Text = "懲戒";
            // 
            // btnExitAll
            // 
            this.btnExitAll.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExitAll.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExitAll.BackColor = System.Drawing.Color.Transparent;
            this.btnExitAll.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExitAll.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnExitAll.Location = new System.Drawing.Point(478, 505);
            this.btnExitAll.Name = "btnExitAll";
            this.btnExitAll.Size = new System.Drawing.Size(103, 23);
            this.btnExitAll.TabIndex = 27;
            this.btnExitAll.Text = "離開";
            this.btnExitAll.Click += new System.EventHandler(this.btnExitAll_Click);
            // 
            // AttendanceUnifyForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(592, 538);
            this.Controls.Add(this.groupPanel4);
            this.Controls.Add(this.lbHelp1);
            this.Controls.Add(this.groupPanel3);
            this.Controls.Add(this.btnExitAll);
            this.DoubleBuffered = true;
            this.MaximizeBox = true;
            this.Name = "AttendanceUnifyForm";
            this.Text = "缺曠學期統計";
            this.Load += new System.EventHandler(this.AttendanceUnifyForm_Load);
            this.groupPanel4.ResumeLayout(false);
            this.groupPanel4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAttendance)).EndInit();
            this.groupPanel3.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.LabelX lbHelp1;
        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel4;
        private DevComponents.DotNetBar.LabelX lbHelp2;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgvAttendance;
        private DevComponents.DotNetBar.ButtonX btnSaveAttendanceStatistics;
        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel3;
        private JHSchool.Behavior.Legacy.ListViewEx listViewAttendance;
        private DevComponents.DotNetBar.ButtonX btnAttendanceNew;
        private DevComponents.DotNetBar.ButtonX btnAttendanceDelete;
        private DevComponents.DotNetBar.ButtonX buttonX4;
        private DevComponents.DotNetBar.ButtonX btnAttendanceEdit;
        private DevComponents.DotNetBar.TabItem tabItem2;
        private DevComponents.DotNetBar.ButtonX btnExitAll;
    }
}