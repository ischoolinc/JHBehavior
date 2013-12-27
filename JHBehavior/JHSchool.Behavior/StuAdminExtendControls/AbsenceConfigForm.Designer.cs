namespace JHSchool.Behavior.StuAdminExtendControls
{
    partial class AbsenceConfigForm
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dataGridView = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.Col1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Col2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Col3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Col5 = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.btnSave = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.btnPrintOut = new DevComponents.DotNetBar.ButtonX();
            this.btnPrintIn = new DevComponents.DotNetBar.ButtonX();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView
            // 
            this.dataGridView.BackgroundColor = System.Drawing.Color.White;
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Col1,
            this.Col2,
            this.Col3,
            this.Col5});
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle6.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView.DefaultCellStyle = dataGridViewCellStyle6;
            this.dataGridView.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dataGridView.Location = new System.Drawing.Point(12, 12);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.RowTemplate.Height = 24;
            this.dataGridView.Size = new System.Drawing.Size(378, 383);
            this.dataGridView.TabIndex = 0;
            this.dataGridView.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_CellEndEdit);
            this.dataGridView.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_CellEnter);
            // 
            // Col1
            // 
            this.Col1.HeaderText = "缺曠名稱";
            this.Col1.Name = "Col1";
            // 
            // Col2
            // 
            this.Col2.HeaderText = "縮寫";
            this.Col2.Name = "Col2";
            this.Col2.Width = 65;
            // 
            // Col3
            // 
            this.Col3.HeaderText = "熱鍵";
            this.Col3.Name = "Col3";
            this.Col3.Width = 65;
            // 
            // Col5
            // 
            this.Col5.HeaderText = "不影響全勤";
            this.Col5.Name = "Col5";
            this.Col5.Width = 95;
            // 
            // btnSave
            // 
            this.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSave.BackColor = System.Drawing.Color.Transparent;
            this.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSave.Location = new System.Drawing.Point(234, 401);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "儲存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(315, 401);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.TabIndex = 2;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnPrintOut
            // 
            this.btnPrintOut.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPrintOut.BackColor = System.Drawing.Color.Transparent;
            this.btnPrintOut.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPrintOut.Location = new System.Drawing.Point(12, 401);
            this.btnPrintOut.Name = "btnPrintOut";
            this.btnPrintOut.Size = new System.Drawing.Size(57, 23);
            this.btnPrintOut.TabIndex = 3;
            this.btnPrintOut.Text = "匯出";
            this.btnPrintOut.Click += new System.EventHandler(this.btnPrintOut_Click);
            // 
            // btnPrintIn
            // 
            this.btnPrintIn.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPrintIn.BackColor = System.Drawing.Color.Transparent;
            this.btnPrintIn.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPrintIn.Location = new System.Drawing.Point(75, 401);
            this.btnPrintIn.Name = "btnPrintIn";
            this.btnPrintIn.Size = new System.Drawing.Size(57, 23);
            this.btnPrintIn.TabIndex = 4;
            this.btnPrintIn.Text = "匯入";
            this.btnPrintIn.Click += new System.EventHandler(this.btnPrintIn_Click);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.FileName = "匯出缺曠類別管理";
            this.saveFileDialog1.Filter = "Excel (*.xls)|*.xls";
            // 
            // AbsenceConfigForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(402, 436);
            this.Controls.Add(this.btnPrintIn);
            this.Controls.Add(this.btnPrintOut);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.dataGridView);
            this.MaximumSize = new System.Drawing.Size(410, 470);
            this.MinimumSize = new System.Drawing.Size(410, 470);
            this.Name = "AbsenceConfigForm";
            this.Text = "缺曠類別管理";
            this.Load += new System.EventHandler(this.AbsenceConfigForm_Load);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.AbsenceConfigForm_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.Controls.DataGridViewX dataGridView;
        private DevComponents.DotNetBar.ButtonX btnSave;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private System.Windows.Forms.DataGridViewTextBoxColumn Col1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Col2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Col3;
        private System.Windows.Forms.DataGridViewCheckBoxColumn Col5;
        private DevComponents.DotNetBar.ButtonX btnPrintOut;
        private DevComponents.DotNetBar.ButtonX btnPrintIn;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    }
}