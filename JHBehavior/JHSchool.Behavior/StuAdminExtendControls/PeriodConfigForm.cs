using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using Framework;
using FISCA.DSAUtil;
using Framework.Feature;
using System.Xml;
using FISCA.LogAgent;
using Aspose.Cells;

namespace JHSchool.Behavior.StuAdminExtendControls
{
    //匯入功能
    //節次名稱 不得重覆

    public partial class PeriodConfigForm : FISCA.Presentation.Controls.BaseForm
    {

        private Dictionary<string, string> DicLogBefor = new Dictionary<string, string>();
        private Dictionary<string, string> DicLogAeft = new Dictionary<string, string>();

        //DataGridView更新檢查
        private ChangeListener DataListener { get; set; }
        private bool DataGridViewDataInChange = false;

        public PeriodConfigForm()
        {
            InitializeComponent();
        }

        private void PeriodConfigForm_Load(object sender, EventArgs e)
        {

            //資料更動檢查
            DataListener = new ChangeListener();
            DataListener.Add(new DataGridViewSource(dataGridView));
            DataListener.StatusChanged += new EventHandler<ChangeEventArgs>(DataListener_StatusChanged);

            //取得Xml結構
            DSResponse dsrsp = Config.GetPeriodList();
            DSXmlHelper helper = dsrsp.GetContent();
            List<PeriodInfo> collection = new List<PeriodInfo>();
            foreach (XmlElement element in helper.GetElements("Period"))
            {
                PeriodInfo info = new PeriodInfo(element);
                collection.Add(info);
            }
            collection.Sort(SortByOrder); //排列順序

            foreach (PeriodInfo info in collection)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridView);

                row.Cells[0].Value = info.Name; //節次名稱
                row.Cells[1].Value = info.Type; //類型
                row.Cells[2].Value = info.Sort.ToString(); //顯示順序
                row.Cells[3].Value = info.Aggregated; //統計權重
                dataGridView.Rows.Add(row);

                if (!DicLogBefor.ContainsKey(info.Name))
                {
                    DicLogBefor.Add(info.Name, "節次名稱「" + info.Name + "」" + "類型「" + info.Type + "」" + "顯示順序「" + info.Sort.ToString() + "」");
                }
            }

            DataListener.Reset();
            DataListener.ResumeListen();
        }

        /// <summary>
        /// 儲存
        /// </summary>
        private void btnSave_Click(object sender, EventArgs e)
        {
            DicLogAeft.Clear(); //Log

            DataListener.SuspendListen(); //終止變更判斷

            //資料檢查
            if (!ValidateRow())
            {
                FISCA.Presentation.Controls.MsgBox.Show("輸入資料有誤，請修正後再行儲存。", "內容錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                return;
            }

            #region 資料收集
            XmlDocument doc = new XmlDocument();
            XmlElement root = doc.CreateElement("Periods");
            doc.AppendChild(root);

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                if (row.IsNewRow) continue;

                XmlElement period = doc.CreateElement("Period");
                root.AppendChild(period);
                period.SetAttribute("Name", ("" + row.Cells[0].Value).Trim());
                period.SetAttribute("Type", ("" + row.Cells[1].Value).Trim());
                period.SetAttribute("Sort", ("" + row.Cells[2].Value).Trim());
                if (!string.IsNullOrEmpty("" + row.Cells[3].Value))
                {
                    period.SetAttribute("Aggregated", ("" + row.Cells[3].Value).Trim());
                }
                else
                {
                    period.SetAttribute("Aggregated", "1");
                }

                //Log
                string InsertLog = "節次名稱「" + row.Cells[0].Value + "」";
                InsertLog += " 類型「" + row.Cells[1].Value + "」";
                InsertLog += "顯示順序「" + row.Cells[2].Value + "」";

                if (!DicLogAeft.ContainsKey("" + row.Cells[0].Value))
                {
                    DicLogAeft.Add("" + row.Cells[0].Value, InsertLog);
                }
            }
            #endregion

            #region 儲存
            string warningMsg = "說明:\n修改節次名稱,建議使用資料合理性檢查\n〔學生缺曠資料與系統節次〕\n以確認系統內存資料的正確性！\n\n是否儲存變更?";
            if (FISCA.Presentation.Controls.MsgBox.Show(warningMsg, "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                return;

            DSXmlHelper helper = new DSXmlHelper("Lists");
            helper.AddElement("List");
            helper.AddElement("List", "Content", root.OuterXml, true);
            helper.AddElement("List", "Condition");
            helper.AddElement("List/Condition", "Name", Config.LIST_PERIODS_NAME);
            //儲存
            try
            {
                Config.Update(new DSRequest(helper));
            }
            catch (Exception exception)
            {
                FISCA.Presentation.Controls.MsgBox.Show("更新失敗 :" + exception.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                Config.Reset(Config.LIST_PERIODS);
            }
            catch
            {
                FISCA.Presentation.Controls.MsgBox.Show("資料重設失敗，新設定值將於下次啟動系統後生效!", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            #endregion

            #region Log
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("「每日節次」已被修改。");
            sb.AppendLine("修改前：");
            foreach (string each in DicLogBefor.Keys)
            {
                sb.AppendLine(DicLogBefor[each]);
            }
            sb.AppendLine("修改後：");
            foreach (string each in DicLogAeft.Keys)
            {
                sb.AppendLine(DicLogAeft[each]);
            }

            ApplicationLog.Log("學務系統.每日節次管理", "修改每日節次", sb.ToString());
            #endregion

            DataGridViewDataInChange = false;
            FISCA.Presentation.Controls.MsgBox.Show("儲存成功!", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);

            this.Close();

        }

        /// <summary>
        /// 離開
        /// </summary>
        private void btnExit_Click(object sender, EventArgs e)
        {
            //if (DataGridViewDataInChange)
            //{
            //    DialogResult dr = FISCA.Presentation.Controls.MsgBox.Show("資料已變更,是否離開?", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2);
            //    if (dr == DialogResult.Yes)
            //    {
            //        this.Close();
            //    }
            //}
            //else
            //{
            this.Close();
            //}
        }

        /// <summary>
        /// 輸入後,檢查是否重覆
        /// </summary>
        private void dataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView.Rows[e.RowIndex].IsNewRow)
                return;

            DataGridViewCell cell = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];

            if (e.ColumnIndex == 0) //節次名稱(不可空白/不可重覆)
            {
                CheckNameRepeat(e.ColumnIndex, e.RowIndex);
            }
            else if (e.ColumnIndex == 1) //類型(不可空白)
            {
                if ("" + cell.Value == string.Empty)
                {
                    cell.ErrorText = "類型不得空白!"; //可重覆
                }
                else
                {
                    cell.ErrorText = ""; //可重覆
                }
            }
            else if (e.ColumnIndex == 2) //顯示順序(不可空白/不可重覆/必須是數字)
            {
                CheckNameRepeat(e.ColumnIndex, e.RowIndex);

                int CellInt;
                if (cell.ErrorText == "")
                {
                    if (!int.TryParse("" + cell.Value, out CellInt))
                    {
                        cell.ErrorText = "必須輸入數字!";
                    }
                    else
                    {
                        cell.ErrorText = "";
                    }
                }
            }
        }

        /// <summary>
        /// 資料重覆檢查
        /// </summary>
        private void CheckNameRepeat(int ColumnIndex, int RowIndex)
        {
            string Name = "" + dataGridView.Rows[RowIndex].Cells[ColumnIndex].Value;
            DataGridViewRow row = dataGridView.Rows[RowIndex];

            List<string> list = new List<string>();

            foreach (DataGridViewRow TalRow in dataGridView.Rows)
            {
                if (TalRow.IsNewRow)
                    continue;

                foreach (DataGridViewCell cell in TalRow.Cells)
                {
                    if (cell.ColumnIndex == ColumnIndex) //同一行
                    {
                        if (cell.RowIndex != RowIndex) //不同列
                        {
                            list.Add("" + cell.Value);
                        }
                    }
                }
            }

            if (Name == string.Empty)
            {
                row.Cells[ColumnIndex].ErrorText = "必須輸入內容!";
            }
            else if (list.Contains(Name))
            {
                row.Cells[ColumnIndex].ErrorText = "資料重覆,請重新輸入!";
            }
            else
            {
                row.Cells[ColumnIndex].ErrorText = "";
            }
        }

        //資料檢查
        private bool ValidateRow()
        {
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                if (row.IsNewRow)
                    continue;

                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (!cell.Visible)
                        continue;

                    if (cell.ErrorText != "")
                    {
                        return false;
                    }
                    else if ("" + cell.Value == "")
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        void DataListener_StatusChanged(object sender, ChangeEventArgs e)
        {
            DataGridViewDataInChange = true;
        }

        private static int SortByOrder(PeriodInfo info1, PeriodInfo info2)
        {
            return info1.Sort.CompareTo(info2.Sort);
        }

        /// <summary>
        /// BeginEdit
        /// </summary>
        private void dataGridView_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                dataGridView.ImeMode = ImeMode.OnHalf;
                dataGridView.ImeMode = ImeMode.Off;
            }

            if (dataGridView.SelectedCells.Count == 1)
                dataGridView.BeginEdit(true);
        }

        private void PeriodConfigForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (DataGridViewDataInChange)
            {
                DialogResult dr = FISCA.Presentation.Controls.MsgBox.Show("資料已變更,是否離開?", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.No)
                {
                    e.Cancel = true;
                }
            }
        }

        private void btnPrintOut_Click(object sender, EventArgs e)
        {
            #region 匯出
            if (saveFileDialog1.ShowDialog() != DialogResult.OK) return;

            DataGridViewExport export = new DataGridViewExport(dataGridView);
            export.Save(saveFileDialog1.FileName);

            if (new CompleteForm().ShowDialog() == DialogResult.Yes)
                System.Diagnostics.Process.Start(saveFileDialog1.FileName);
            #endregion
        }

        private void btnPrintIn_Click(object sender, EventArgs e)
        {

            #region 確認畫面
            DialogResult dr = FISCA.Presentation.Controls.MsgBox.Show("匯入每日節次設定值\n將完全覆蓋目前之資料狀態\n(建議可將原資料匯出備份)\n\n請確認繼續?", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2);
            if (dr != DialogResult.Yes)
                return;


            Workbook wb = new Workbook();

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "選擇要匯入的每日節次設定值";
            ofd.Filter = "Excel檔案 (*.xls)|*.xls";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    wb.Open(ofd.FileName);
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "開啟檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
                return;

            //必要欄位
            List<string> requiredHeaders = new List<string>(new string[] { "節次名稱", "類型", "顯示順序" });
            //欄位標題的索引
            Dictionary<string, int> headers = new Dictionary<string, int>();
            Worksheet ws = wb.Worksheets[0];
            for (int i = 0; i <= ws.Cells.MaxDataColumn; i++)
            {
                string header = ws.Cells[0, i].StringValue;
                if (requiredHeaders.Contains(header))
                    headers.Add(header, i);
            }

            //如果使用者匯入檔的欄位與必要欄位不符，則停止匯入
            if (headers.Count != requiredHeaders.Count)
            {
                StringBuilder builder = new StringBuilder(string.Empty);
                builder.AppendLine("匯入格式不符合。");
                builder.AppendLine("匯入資料標題必須包含：");
                builder.AppendLine(string.Join(",", requiredHeaders.ToArray()));
                FISCA.Presentation.Controls.MsgBox.Show(builder.ToString());
                return;
            }

            //string CellName1 = wb.Worksheets[0].Cells[0, 0].StringValue;
            //string CellName2 = wb.Worksheets[0].Cells[0, 1].StringValue;
            //string CellName3 = wb.Worksheets[0].Cells[0, 2].StringValue;
            //string CellName4 = wb.Worksheets[0].Cells[0, 3].StringValue;

            //if (CellName1 != "節次名稱" || CellName2 != "類型" || CellName3 != "顯示順序" || CellName4 != "統計權重")
            //{
            //    FISCA.Presentation.Controls.MsgBox.Show("匯入格式不符合。\n匯入資料標題必須依照:\n節次名稱,類型,顯示順序,統計權重\n目前為:\n" + CellName1 + "," + CellName2 + "," + CellName3 + "," + CellName4);
            //    return;
            //}
            #endregion

            #region 匯入

            XmlDocument doc = new XmlDocument();
            XmlElement root = doc.CreateElement("Periods");
            doc.AppendChild(root);

            #region 匯入重覆問題
            List<string> NameList1 = new List<string>();
            StringBuilder NameSb = new StringBuilder();
            for (int x = 1; x <= wb.Worksheets[0].Cells.MaxDataRow; x++) //每一Row
            {
                string name = ws.Cells[x, headers["節次名稱"]].StringValue;

                if (string.IsNullOrEmpty(name.Trim())) //沒有缺曠名稱則跳過
                    continue;

                if (!NameList1.Contains(name.Trim()))
                {
                    NameList1.Add(name);
                }
                else
                {
                    NameSb.AppendLine("節次名稱重覆:" + name);
                }

            }
            if (!string.IsNullOrEmpty(NameSb.ToString()))
            {
                FISCA.Presentation.Controls.MsgBox.Show("匯入每日節次發生錯誤:\n" + NameSb.ToString());
                return;
            }
            #endregion

            for (int x = 1; x <= wb.Worksheets[0].Cells.MaxDataRow; x++) //每一Row
            {
                string name = ws.Cells[x, headers["節次名稱"]].StringValue;
                string type = ws.Cells[x, headers["類型"]].StringValue;
                string sort = ws.Cells[x, headers["顯示順序"]].StringValue;

                if (string.IsNullOrEmpty(name.Trim())) //沒有節次名稱則跳過
                    continue;

                XmlElement period = doc.CreateElement("Period");
                root.AppendChild(period);

                period.SetAttribute("Name", name.Trim());
                if (!string.IsNullOrEmpty(type.Trim()))
                {
                    period.SetAttribute("Type", type.Trim());
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("匯入失敗,類型不可為空白!!\n此錯誤發生於節次名稱:[" + name + "]");
                    return;
                }
                if (CheckInt(sort.Trim()))
                    period.SetAttribute("Sort", sort.Trim());
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("匯入失敗,顯示順序必須是數字!!\n此錯誤發生於節次名稱:[" + name + "]");
                    return;
                }

                period.SetAttribute("Aggregated", "1"); //統計權重目前沒有用到，預設為 1

                //if (CheckDouble(aggregated))
                //    period.SetAttribute("Aggregated", aggregated);
                //else
                //{
                //    FISCA.Presentation.Controls.MsgBox.Show("匯入失敗,統計權重必須是數字(可包含小數點)!");
                //    return;
                //}
            }

            DSXmlHelper helper = new DSXmlHelper("Lists");
            helper.AddElement("List");
            helper.AddElement("List", "Content", root.OuterXml, true);
            helper.AddElement("List", "Condition");
            helper.AddElement("List/Condition", "Name", Config.LIST_PERIODS_NAME);

            //儲存
            try
            {
                Config.Update(new DSRequest(helper));
            }
            catch (Exception exception)
            {
                FISCA.Presentation.Controls.MsgBox.Show("更新失敗 :" + exception.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                Config.Reset(Config.LIST_PERIODS);
            }
            catch
            {
                FISCA.Presentation.Controls.MsgBox.Show("資料重設失敗!", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }



            ApplicationLog.Log("學務系統.每日節次設定檔", "匯入每日節次設定檔", "每日節次設定檔，已將每日節次管理內容覆蓋匯入。");
            FISCA.Presentation.Controls.MsgBox.Show("匯入成功!\n新設定將於畫面重新開啟時生效!", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            DataGridViewDataInChange = false;
            this.Close();

            #endregion
        }

        private bool CheckInt(string j)
        {
            int HotKeyint;
            if (int.TryParse(j, out HotKeyint))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private bool CheckDouble(string k)
        {
            Double HotKeyint;
            if (Double.TryParse(k, out HotKeyint))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
