using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace JHSchool.Behavior.Legacy
{
    public class ConflictDialog : BaseForm
    {
        private DevComponents.DotNetBar.Controls.TextBoxX textBoxX1;
        private DevComponents.DotNetBar.ButtonX buttonX1;
        private DevComponents.DotNetBar.ButtonX buttonX2;
        private readonly List<ConflictInfo> _conflicts;

        public ConflictDialog(List<ConflictInfo> conflicts)
        {
            _conflicts = conflicts;
            InitializeComponent();
            BuildContent();
        }

        private void BuildContent()
        {
            Text = "資料衝突警告";
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;

            buttonX1.DialogResult = DialogResult.Yes;
            buttonX2.DialogResult = DialogResult.No;
            AcceptButton = buttonX1;
            CancelButton = buttonX2;

            textBoxX1.ReadOnly = true;
            textBoxX1.Font = new Font("微軟正黑體", 9.5f);

            StringBuilder sb = new StringBuilder();
            sb.AppendLine(string.Format("有 {0} 筆缺曠資料在您編輯期間已被他人修改：", _conflicts.Count));
            sb.AppendLine();

            foreach (ConflictInfo c in _conflicts)
            {
                sb.AppendLine(string.Format("{0} {1}號 {2}　{3}",
                    c.ClassName, c.SeatNo, c.Name, c.OccurDate.ToShortDateString()));

                if (c.DeletedByOther)
                {
                    sb.AppendLine("  → 此筆缺曠紀錄已被他人刪除");
                }
                else
                {
                    sb.AppendLine(string.Format("  {0,-8}  {1,-8}  {2,-8}  {3}", "節次", "載入時", "您的修改", "他人改後"));
                    sb.AppendLine("  " + new string('-', 44));
                    foreach (PeriodDiff d in c.PeriodDiffs)
                    {
                        string before  = string.IsNullOrEmpty(d.BeforeAbsence)  ? "（無）"    : d.BeforeAbsence;
                        string user    = string.IsNullOrEmpty(d.UserAbsence)    ? "（清除）"  : d.UserAbsence;
                        string current = string.IsNullOrEmpty(d.CurrentAbsence) ? "（已清除）": d.CurrentAbsence;
                        sb.AppendLine(string.Format("  {0,-8}  {1,-8}  {2,-8}  {3}", d.PeriodName, before, user, current));
                    }
                }

                sb.AppendLine();
            }

            textBoxX1.Text = sb.ToString();
        }

        private void InitializeComponent()
        {
            this.textBoxX1 = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.buttonX1 = new DevComponents.DotNetBar.ButtonX();
            this.buttonX2 = new DevComponents.DotNetBar.ButtonX();
            this.SuspendLayout();
            // 
            // textBoxX1
            // 
            this.textBoxX1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            // 
            // 
            // 
            this.textBoxX1.Border.Class = "TextBoxBorder";
            this.textBoxX1.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.textBoxX1.Location = new System.Drawing.Point(12, 8);
            this.textBoxX1.Multiline = true;
            this.textBoxX1.Name = "textBoxX1";
            this.textBoxX1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxX1.Size = new System.Drawing.Size(760, 512);
            this.textBoxX1.TabIndex = 2;
            // 
            // buttonX1
            // 
            this.buttonX1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.buttonX1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonX1.BackColor = System.Drawing.Color.Transparent;
            this.buttonX1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.buttonX1.Location = new System.Drawing.Point(527, 531);
            this.buttonX1.Name = "buttonX1";
            this.buttonX1.Size = new System.Drawing.Size(164, 23);
            this.buttonX1.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.buttonX1.TabIndex = 3;
            this.buttonX1.Text = "仍要覆蓋儲存";
            // 
            // buttonX2
            // 
            this.buttonX2.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.buttonX2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonX2.BackColor = System.Drawing.Color.Transparent;
            this.buttonX2.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.buttonX2.Location = new System.Drawing.Point(697, 531);
            this.buttonX2.Name = "buttonX2";
            this.buttonX2.Size = new System.Drawing.Size(75, 23);
            this.buttonX2.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.buttonX2.TabIndex = 4;
            this.buttonX2.Text = "取消操作";
            // 
            // ConflictDialog
            // 
            this.ClientSize = new System.Drawing.Size(784, 561);
            this.Controls.Add(this.buttonX2);
            this.Controls.Add(this.buttonX1);
            this.Controls.Add(this.textBoxX1);
            this.DoubleBuffered = true;
            this.Name = "ConflictDialog";
            this.ResumeLayout(false);

        }
    }

    public class ConflictInfo
    {
        public string StudentID { get; set; }
        public string ClassName { get; set; }
        public string SeatNo { get; set; }
        public string StudentNumber { get; set; }
        public string Name { get; set; }
        public DateTime OccurDate { get; set; }
        public bool DeletedByOther { get; set; }
        public List<PeriodDiff> PeriodDiffs { get; set; }
        public ConflictInfo() { PeriodDiffs = new List<PeriodDiff>(); }
    }

    public class PeriodDiff
    {
        public string PeriodName { get; set; }
        public string BeforeAbsence { get; set; }
        public string UserAbsence { get; set; }
        public string CurrentAbsence { get; set; }
    }
}
