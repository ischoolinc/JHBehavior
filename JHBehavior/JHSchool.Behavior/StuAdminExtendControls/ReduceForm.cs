using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
//using SmartSchool.Common;
using FISCA.DSAUtil;
using System.Xml;
using DevComponents.DotNetBar.Controls;
using DevComponents.DotNetBar;
using Framework.Feature;
using Framework;
using FISCA.LogAgent;

namespace JHSchool.Behavior.StuAdminExtendControls
{
    public partial class ReduceForm : FISCA.Presentation.Controls.BaseForm
    {
        StringBuilder sb = new StringBuilder();

        public ReduceForm()
        {
            InitializeComponent();
        }

        private void ReduceForm_Load(object sender, EventArgs e)
        {
            DSResponse dsrsp = Config.GetMDReduce();
            if (!dsrsp.HasContent)
            {
                FISCA.Presentation.Controls.MsgBox.Show("取得對照表失敗 : " + dsrsp.GetFault().Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            DSXmlHelper helper = dsrsp.GetContent();
            txtMAB.Text = helper.GetText("Merit/AB");
            txtMBC.Text = helper.GetText("Merit/BC");
            txtDAB.Text = helper.GetText("Demerit/AB");
            txtDBC.Text = helper.GetText("Demerit/BC");

            sb.AppendLine("「功過換算表」已被修改。");
            sb.AppendLine("修改前：");
            sb.AppendLine("「1大功」等於「" + txtMAB.Text + "小功」");
            sb.AppendLine("「1小功」等於「" + txtMBC.Text + "嘉獎」");
            sb.AppendLine("「1大過」等於「" + txtDAB.Text + "小過」");
            sb.AppendLine("「1小過」等於「" + txtDBC.Text + "嘉獎」");
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!IsValid()) return;
            //獎懲單位換算表
            XmlDocument doc = new XmlDocument();
            XmlElement root = doc.CreateElement("Reduce");
            doc.AppendChild(root);

            XmlElement element = doc.CreateElement("Merit");
            root.AppendChild(element);
            XmlElement ab = doc.CreateElement("AB");
            element.AppendChild(ab);
            ab.InnerText = txtMAB.Text;
            XmlElement bc = doc.CreateElement("BC");
            element.AppendChild(bc);
            bc.InnerText = txtMBC.Text;

            element = doc.CreateElement("Demerit");
            root.AppendChild(element);
            ab = doc.CreateElement("AB");
            element.AppendChild(ab);
            ab.InnerText = txtDAB.Text;
            bc = doc.CreateElement("BC");
            element.AppendChild(bc);
            bc.InnerText = txtDBC.Text;

            sb.AppendLine("修改後：");
            sb.AppendLine("「1大功」等於「" + txtMAB.Text + "小功」");
            sb.AppendLine("「1小功」等於「" + txtMBC.Text + "嘉獎」");
            sb.AppendLine("「1大過」等於「" + txtDAB.Text + "小過」");
            sb.AppendLine("「1小過」等於「" + txtDBC.Text + "嘉獎」");

            try
            {
                DSXmlHelper helper = new DSXmlHelper("Lists");
                helper.AddElement("List");
                helper.AddElement("List", "Content");
                helper.AddElement("List/Content", doc.DocumentElement);
                helper.AddElement("List", "Condition");
                helper.AddElement("List/Condition", "Name", "獎懲單位換算表");
                Config.Update(new DSRequest(helper));
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("儲存失敗 : " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            ApplicationLog.Log("學務系統.功過換算管理", "修改功過換算", sb.ToString());
            MsgBox.Show("儲存成功!", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }

        private bool IsValid()
        {
            error.Clear();
            error.Tag = true;
            ValidInt(txtMAB, lblMAB);
            ValidInt(txtMBC, lblMBC);
            ValidInt(txtDAB, lblDAB);
            ValidInt(txtDBC, lblDBC);
            return bool.Parse(error.Tag.ToString());
        }

        private void ValidInt(TextBoxX txt, LabelX lbl)
        {
            int i;
            if (!int.TryParse(txt.Text, out i))
            {
                error.Tag = false;
                error.SetError(lbl, "必須為數字");
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}