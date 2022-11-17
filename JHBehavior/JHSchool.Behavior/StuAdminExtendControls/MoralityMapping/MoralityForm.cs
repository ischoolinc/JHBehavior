using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Xml;
using System.Windows.Forms;
//using SmartSchool.Common;
//using SmartSchool.Feature.Basic;
using FISCA.DSAUtil;
using Aspose.Cells;
using System.IO;
using Framework.Feature;
using Framework;
using FISCA.Presentation.Controls;
using FISCA.LogAgent;

namespace JHSchool.Behavior.StuAdminExtendControls.MoralityMapping
{
    public partial class MoralityForm : FISCA.Presentation.Controls.BaseForm
    {
        private Dictionary<string, string> _origList = new Dictionary<string, string>();
        private bool _isSave = true;
        private int _SelectedRowIndex;

        public MoralityForm()
        {
            InitializeComponent();
            InitialList();

            List<string> cols = new List<string>() { "���y�N�X" };
            Campus.Windows.DataGridViewImeDecorator dec = new Campus.Windows.DataGridViewImeDecorator(this.dataGridViewX1, cols);

        }

        private void InitialList()
        {
            DSResponse dsrsp = Config.GetMoralCommentCodeList();
            foreach (XmlElement var in dsrsp.GetContent().GetElements("Morality"))
            {
                int index = dataGridViewX1.Rows.Add();
                DataGridViewRow row = dataGridViewX1.Rows[index];
                row.Cells[Code.Name].Value = var.GetAttribute("Code");
                row.Cells[Comment.Name].Value = var.GetAttribute("Comment");
                _origList.Add(var.GetAttribute("Code"), var.GetAttribute("Comment"));
            }
        }

        private bool ValidateList()
        {
            dataGridViewX1.EndEdit();
            bool valid = true;
            _isSave = true;
            List<string> codeList = new List<string>();

            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                if (row.IsNewRow) continue;
                row.Cells[Code.Index].ErrorText = "";
                if (row.Cells[Code.Index].Value == null)
                {
                    row.Cells[Code.Index].ErrorText = "���ର�ť�";
                    valid = false;
                    break;
                }

                string codeValue = row.Cells[Code.Name].Value.ToString();

                if (!codeList.Contains(codeValue))
                    codeList.Add(codeValue);
                else
                {
                    row.Cells[Code.Name].ErrorText = "�W�٭���";
                    valid = false;
                    break;
                }

                //�ˬd��ƬO�_�ܰ�
                if (_isSave)
                {
                    if (_origList.ContainsKey(codeValue))
                    {
                        if (_origList[codeValue] != ((row.Cells[Comment.Name].Value != null) ? row.Cells[Comment.Name].Value.ToString() : ""))
                            _isSave = false;
                    }
                    else
                        _isSave = false;
                }
            }

            if (_isSave)
            {
                if (dataGridViewX1.Rows.Count - 1 != _origList.Keys.Count)
                    _isSave = false;
            }

            return valid;
        }

        private bool Save()
        {
            XmlDocument doc = new XmlDocument();
            XmlElement content = doc.CreateElement("Content");

            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                if (row.IsNewRow) continue;
                XmlElement morality = doc.CreateElement("Morality");
                morality.SetAttribute("Code", row.Cells[Code.Name].Value.ToString());
                morality.SetAttribute("Comment", (row.Cells[Comment.Name].Value != null) ? row.Cells[Comment.Name].Value.ToString() : "");
                content.AppendChild(morality);

                if (!_origList.ContainsKey(row.Cells[Code.Name].Value.ToString()))
                    _origList.Add(row.Cells[Code.Name].Value.ToString(), (row.Cells[Comment.Name].Value != null) ? row.Cells[Comment.Name].Value.ToString() : "");
            }

            try
            {
                Config.SetMoralCommentCodeList(content);
                _isSave = true;
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!ValidateList())
            {
                FISCA.Presentation.Controls.MsgBox.Show("��Ʀ��~�A�|���x�s");
                return;
            }

            if (Save())
            {
                FISCA.Presentation.Controls.MsgBox.Show("�x�s���\�C");
                ApplicationLog.Log("�ǰȨt��.�ɮv���y�N�X��", "�ק�ɮv���y�N�X", "�u�ɮv���y�N�X��v�w�Q�ק�C");
                this.Close();
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("�x�s���ѡC");
                return;
            }

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            Workbook wb = new Workbook();
            Dictionary<string, string> importCodeList = new Dictionary<string, string>();

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "��ܭn�פJ���ɮv���y�N�X��";
            ofd.Filter = "Excel�ɮ� (*.xlsx)|*.xlsx";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    wb.Open(ofd.FileName);
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("���w���|�L�k�s���C", "�}���ɮץ���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
                return;

            if (wb.Worksheets[0].Cells[0, 0].StringValue != "���y�N�X" || wb.Worksheets[0].Cells[0, 1].StringValue != "���y���e")
            {
                FISCA.Presentation.Controls.MsgBox.Show("�פJ�榡���ŦX�C");
                return;
            }

            ImportConfirm form = new ImportConfirm();
            if (form.ShowDialog() == DialogResult.OK)
            {
                int rowIndex = 1;
                Worksheet ws = wb.Worksheets[0];

                while (!string.IsNullOrEmpty(ws.Cells[rowIndex, 0].StringValue))
                {
                    string code = ws.Cells[rowIndex, 0].StringValue;
                    string comment = ws.Cells[rowIndex, 1].StringValue;

                    if (!importCodeList.ContainsKey(code))
                        importCodeList.Add(code, comment);
                    else
                        importCodeList[code] = comment;
                    rowIndex++;
                }

                if (form.Overwrite)
                {
                    dataGridViewX1.Rows.Clear();
                    foreach (string key in importCodeList.Keys)
                    {
                        int index = dataGridViewX1.Rows.Add();
                        DataGridViewRow row = dataGridViewX1.Rows[index];
                        row.Cells[Code.Name].Value = key;
                        row.Cells[Comment.Name].Value = importCodeList[key];
                    }
                    ApplicationLog.Log("�ǰȨt��.�ɮv���y�N�X��", "�פJ�ɮv���y�N�X", "�u�ɮv���y�N�X��v�w�Q�i��פJ�s�W�ާ@�C");
                }
                else
                {
                    Dictionary<string, int> OriginalCodeListIndex = new Dictionary<string, int>();
                    List<int> delete = new List<int>();

                    foreach (DataGridViewRow row in dataGridViewX1.Rows)
                    {
                        if (row.IsNewRow) continue;
                        if (row.Cells[Code.Name].Value != null)
                        {
                            string code = row.Cells[Code.Name].Value.ToString();
                            if (!OriginalCodeListIndex.ContainsKey(code))
                                OriginalCodeListIndex.Add(code, row.Index);
                            else
                            {
                                delete.Add(OriginalCodeListIndex[code]);
                                OriginalCodeListIndex[code] = row.Index;
                            }
                        }
                    }

                    foreach (string key in importCodeList.Keys)
                    {
                        if (OriginalCodeListIndex.ContainsKey(key))
                            dataGridViewX1.Rows[OriginalCodeListIndex[key]].Cells[Comment.Name].Value = importCodeList[key];
                        else
                        {
                            int index = dataGridViewX1.Rows.Add();
                            DataGridViewRow row = dataGridViewX1.Rows[index];
                            row.Cells[Code.Name].Value = key;
                            row.Cells[Comment.Name].Value = importCodeList[key];
                        }
                    }

                    foreach (int var in delete)
                    {
                        dataGridViewX1.Rows.RemoveAt(var);
                    }
                    ApplicationLog.Log("�ǰȨt��.�ɮv���y�N�X��", "�פJ�ɮv���y�N�X", "�u�ɮv���y�N�X��v�w�Q�i��פJ�л\�ާ@�C");
                }

                //FISCA.Presentation.Controls.MsgBox.Show("�פJ���\�C\n�`�N�G�t�Ω|���N����x�s�A�бz��ܡy�T�w�z��A�ܵ��y�N�X���T�{�פJ��ƫ����y�x�s�z�C");
                ValidateList();
                FISCA.Presentation.Controls.MsgBox.Show("�w�פJ����!\n���I���x�s�����}�C");
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (!ValidateList())
                return;

            Workbook wb = new Workbook();
            wb.Worksheets.Clear();
            Worksheet ws = wb.Worksheets[wb.Worksheets.Add()];
            ws.Name = "�ɮv���y�N�X��";

            ws.Cells.CreateRange(0, 1, true).ColumnWidth = 10;
            ws.Cells.CreateRange(1, 1, true).ColumnWidth = 40;

            ws.Cells[0, 0].PutValue("���y�N�X");
            ws.Cells[0, 1].PutValue("���y���e");

            int rowIndex = 1;

            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                if (row.IsNewRow) continue;

                ws.Cells[rowIndex, 0].PutValue(row.Cells[Code.Name].Value.ToString());
                ws.Cells[rowIndex, 1].PutValue((row.Cells[Comment.Name].Value != null) ? row.Cells[Comment.Name].Value.ToString() : "");
                rowIndex++;
            }

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "�t�s�s��";
            sfd.FileName = "�ɮv���y�N�X��.xlsx";
            sfd.Filter = "Excel�ɮ� (*.xlsx)|*.xlsx|�Ҧ��ɮ� (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    wb.Save(sfd.FileName);
                    FISCA.Presentation.Controls.MsgBox.Show("�ץX���\�C");
                    ApplicationLog.Log("�ǰȨt��.�ɮv���y�N�X��", "�ץX�ɮv���y�N�X", "�u�ɮv���y�N�X��v�w�Q�ץX�C");
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("���w���|�L�k�s���C", "�t�s�ɮץ���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void MoralityForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!_isSave)
            {
                if (FISCA.Presentation.Controls.MsgBox.Show("��Ʃ|���x�s�A�z�T�w�n���}�H", "", MessageBoxButtons.YesNo) == DialogResult.No)
                    e.Cancel = true;
            }
        }

        private void dataGridViewX1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = true;
        }

        private void dataGridViewX1_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            ValidateList();
        }

        private void dataGridViewX1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && dataGridViewX1.SelectedCells.Count == 1)
                dataGridViewX1.BeginEdit(true);
        }

        private void dataGridViewX1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dataGridViewX1.EndEdit();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            dataGridViewX1.Rows.Insert(_SelectedRowIndex, new DataGridViewRow());
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (_SelectedRowIndex >= 0 && dataGridViewX1.Rows.Count - 1 > _SelectedRowIndex)
                dataGridViewX1.Rows.RemoveAt(_SelectedRowIndex);
        }

        private void dataGridViewX1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex < 0 && e.Button == MouseButtons.Right)
            {
                dataGridViewX1.EndEdit();
                _SelectedRowIndex = e.RowIndex;
                foreach (DataGridViewRow var in dataGridViewX1.SelectedRows)
                {
                    if (var.Index != _SelectedRowIndex)
                        var.Selected = false;
                }
                dataGridViewX1.Rows[_SelectedRowIndex].Selected = true;
                contextMenuStrip1.Show(dataGridViewX1, dataGridViewX1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true).Location);
            }
        }
    }
}