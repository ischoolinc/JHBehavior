using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using FISCA.DSAUtil;
using System.Xml;
using Framework.Feature;
using FISCA.LogAgent;
using Framework;
using Aspose.Cells;
using System.Text.RegularExpressions;

namespace JHSchool.Behavior.StuAdminExtendControls
{
    //匯入功能
    //缺曠名稱 不得重覆
    //縮寫 不得重覆
    //熱鍵 不得重覆

    public partial class AbsenceConfigForm : FISCA.Presentation.Controls.BaseForm
    {
        //修改前Log狀態
        private Dictionary<string, string> DicLogBefor = new Dictionary<string, string>();
        //修改後Log狀態
        private Dictionary<string, string> DicLogAeft = new Dictionary<string, string>();

        //DataGridView更新檢查
        private ChangeListener DataListener { get; set; } 
        private bool DataGridViewDataInChange = false;

        private Regex Pattern { get; set; }

        public AbsenceConfigForm()
        {
            InitializeComponent();

            Pattern = new Regex("^[A-Za-z0-9]+$");
        }

        /// <summary>
        /// 畫面載入
        /// </summary>
        private void AbsenceConfigForm_Load(object sender, EventArgs e)
        {

            //資料更動檢查
            DataListener = new ChangeListener();
            DataListener.Add(new DataGridViewSource(dataGridView));
            DataListener.StatusChanged += new EventHandler<ChangeEventArgs>(DataListener_StatusChanged);

            //取得Xml結構
            DSResponse dsrsp = Config.GetAbsenceList();
            DSXmlHelper helper = dsrsp.GetContent();
            foreach (XmlElement element in helper.GetElements("Absence"))
            {
                //建立缺曠物件
                AbsenceInfo info = new AbsenceInfo(element);

                //建立Row
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridView);
                row.Cells[0].Value = info.Name; //缺曠名稱
                row.Cells[1].Value = info.Abbreviation; //缺曠名稱
                row.Cells[2].Value = info.Hotkey; //缺曠名稱
                row.Cells[3].Value = info.Noabsence; //缺曠名稱
                dataGridView.Rows.Add(row);

                //Log
                if (!DicLogBefor.ContainsKey(info.Name))
                {
                    if (info.Noabsence)
                    {
                        DicLogBefor.Add(info.Name, "缺曠名稱「" + info.Name + "」縮寫「" + info.Abbreviation + "」熱鍵「" + info.Hotkey + "」不影響全勤「是」");
                    }
                    else
                    {
                        DicLogBefor.Add(info.Name, "缺曠名稱「" + info.Name + "」縮寫「" + info.Abbreviation + "」熱鍵「" + info.Hotkey + "」不影響全勤「否」");
                    }
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
            XmlElement root = doc.CreateElement("AbsenceList");
            doc.AppendChild(root);

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                if (row.IsNewRow)
                    continue;

                XmlElement absence = doc.CreateElement("Absence");
                root.AppendChild(absence);
                absence.SetAttribute("Name", ("" + row.Cells[0].Value).Trim());
                absence.SetAttribute("Abbreviation", ("" + row.Cells[1].Value).Trim());
                absence.SetAttribute("HotKey", ("" + row.Cells[2].Value).Trim());
                absence.SetAttribute("Noabsence", ("" + row.Cells[3].Value).Trim());

                if (!DicLogAeft.ContainsKey("" + row.Cells[0].Value))
                {
                    if ("" + row.Cells[3].Value == "True")
                    {
                        DicLogAeft.Add("" + row.Cells[0].Value, "缺曠名稱「" + row.Cells[0].Value + "」縮寫「" + row.Cells[1].Value + "」熱鍵「" + row.Cells[2].Value + "」不影響全勤「是」");
                    }
                    else
                    {
                        DicLogAeft.Add("" + row.Cells[0].Value, "缺曠名稱「" + row.Cells[0].Value + "」縮寫「" + row.Cells[1].Value + "」熱鍵「" + row.Cells[2].Value + "」不影響全勤「否」");
                    }
                }
            } 
            #endregion

            #region 儲存

            string warningMsg = "說明:\n修改節次名稱,建議使用資料合理性檢查\n〔學生缺曠資料與系統假別〕\n以確認系統內存資料的正確性！\n\n是否儲存變更?";
            if (FISCA.Presentation.Controls.MsgBox.Show(warningMsg, "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                return;

            DSXmlHelper helper = new DSXmlHelper("Lists");
            helper.AddElement("List");
            helper.AddElement("List", "Content", root.OuterXml, true);
            helper.AddElement("List", "Condition");
            helper.AddElement("List/Condition", "Name", Config.LIST_ABSENCE_NAME);

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

            //更新
            try
            {
                Config.Reset(Config.LIST_ABSENCE);
            }
            catch
            {
                FISCA.Presentation.Controls.MsgBox.Show("資料重設失敗，新設定值將於下次啟動系統後生效!", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } 
            #endregion

            #region Log
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("「缺曠類別」已被修改。");
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
            ApplicationLog.Log("學務系統.缺曠類別管理", "修改缺曠類別", sb.ToString()); 
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

        private bool ValidateRow()
        {
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                if (row.IsNewRow)
                    continue;

                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.ErrorText != "")
                    {
                        return false;
                    }
                    else if ("" + cell.Value == "" && cell.ColumnIndex != 3)
                    {
                        return false;
                    }
                }
            }

            return true;

            #region 註解掉
            //    bool pass = true;
            //    //如果是新行
            //    if (row.IsNewRow)
            //        return true;
            //    #region 不允許空白
            //    foreach (DataGridViewCell cell in row.Cells)
            //    {
            //        //如果是CheckBox則略過
            //        if (cell is System.Windows.Forms.DataGridViewCheckBoxCell)
            //            continue;                
            //        if ("" + cell.Value == "") //如果Cell是空的
            //        {
            //            cell.ErrorText = "不允許空白";
            //            pass &= false;
            //        }
            //        else if (cell.ErrorText == "不允許空白") //如果Cell已有錯誤訊息 & Cell不是空的 則清空
            //        {
            //            cell.ErrorText = "";
            //        }
            //    }
            //    #endregion
            //    #region 不得重複(名稱　縮寫　熱鍵)
            //    //foreach (DataGridViewRow r in dataGridView.Rows)
            //    //{
            //    //    if (r != row)
            //    //    {
            //    //        foreach (int index in new int[] { colName.Index, colHotKey.Index, colAbbreviation.Index })
            //    //        {
            //    //            if ("" + r.Cells[index].Value == "" + row.Cells[index].Value)
            //    //            {
            //    //                row.Cells[index].ErrorText = "不得重複";
            //    //                dataGridView.UpdateCellErrorText(index, row.Index);
            //    //                pass &= false;
            //    //            }
            //    //            else if (row.Cells[index].ErrorText == "不得重複")
            //    //            {
            //    //                row.Cells[index].ErrorText = "";
            //    //                dataGridView.UpdateCellErrorText(index, row.Index);
            //    //            }
            //    //        }
            //    //    }
            //    //}
            //    #endregion
            //    return pass; 
            #endregion
        }

        /// <summary>
        /// 資料變更的變數
        /// </summary>
        void DataListener_StatusChanged(object sender, ChangeEventArgs e)
        {
            DataGridViewDataInChange = true;
        }

        /// <summary>
        /// 輸入後,檢查是否重覆
        /// </summary>
        private void dataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView.Rows[e.RowIndex].IsNewRow)
                return;

            if (e.ColumnIndex == 0) //缺曠名稱
            {
                CheckNameRepeat(e.ColumnIndex, e.RowIndex);
                //縮寫未填,則填入缺曠名稱第一個字元
                if (string.IsNullOrEmpty("" + dataGridView.Rows[e.RowIndex].Cells[Col2.Index].Value))
                {
                    string col1Value = "" + dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                    if (!string.IsNullOrEmpty(col1Value))
                    {
                        dataGridView.Rows[e.RowIndex].Cells[Col2.Index].Value = col1Value[0];
                        CheckNameRepeat(Col2.Index, e.RowIndex);
                    }
                    //dataGridView.Rows[e.RowIndex].Cells[Col2.Index].Value
                }
            }
            else if (e.ColumnIndex == 1) //縮寫
            {
                CheckNameRepeat(e.ColumnIndex, e.RowIndex);
                if (dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText == "必須輸入內容!")
                {
                    string col1Value = "" + dataGridView.Rows[e.RowIndex].Cells[Col1.Index].Value;
                    if (!string.IsNullOrEmpty(col1Value))
                    {
                        dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = col1Value[0];
                        CheckNameRepeat(e.ColumnIndex, e.RowIndex);
                    }
                }
            }
            else if (e.ColumnIndex == 2) //熱鍵
            {
                CheckNameRepeat(e.ColumnIndex, e.RowIndex);
            }
        }

        /// <summary>
        /// 資料重覆檢查,空值檢查
        /// </summary>
        /// <param name="ColumnIndex">傳入Column的Index</param>
        /// <param name="row">傳入Row的Index</param>
        private void CheckNameRepeat(int ColumnIndex,int RowIndex)
        {
            string Name = ""+dataGridView.Rows[RowIndex].Cells[ColumnIndex].Value;
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

        private void AbsenceConfigForm_FormClosing(object sender, FormClosingEventArgs e)
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
            DialogResult dr = FISCA.Presentation.Controls.MsgBox.Show("匯入缺曠類別設定值\n將完全覆蓋目前之資料狀態\n(建議可將原資料匯出備份)\n\n請確認繼續?", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2);
            if (dr != DialogResult.Yes)
                return;

            Workbook wb = new Workbook();

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "選擇要匯入的缺曠類別設定值";
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
            List<string> requiredHeaders = new List<string>(new string[] { "缺曠名稱", "縮寫", "熱鍵", "不影響全勤" });
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

            //if (CellName1 != "缺曠名稱" || CellName2 != "縮寫" || CellName3 != "熱鍵" || CellName4 != "不影響全勤")
            //{
            //    FISCA.Presentation.Controls.MsgBox.Show("匯入格式不符合。\n匯入資料標題必須依照:\n缺曠名稱,縮寫,熱鍵,不影響全勤\n目前為:\n" + CellName1 + "," + CellName2 + "," + CellName3 + "," + CellName4);
            //    return;
            //}

            #endregion

            #region 匯入

            XmlDocument doc = new XmlDocument();
            XmlElement root = doc.CreateElement("AbsenceList");
            doc.AppendChild(root);

            #region 匯入重覆問題
            List<string> NameList1 = new List<string>();
            List<string> NameList2 = new List<string>();
            List<string> NameList3 = new List<string>();
            StringBuilder NameSb = new StringBuilder();
            for (int x = 1; x <= wb.Worksheets[0].Cells.MaxDataRow; x++) //每一Row
            {
                string name = ws.Cells[x, headers["缺曠名稱"]].StringValue;

                if (string.IsNullOrEmpty(name.Trim())) //沒有缺曠名稱則跳過
                    continue;

                if (!NameList1.Contains(name.Trim()))
                {
                    NameList1.Add(name);
                }
                else
                {
                    NameSb.AppendLine("缺曠名稱重覆:" + name);
                }
                string abbreviation = ws.Cells[x, headers["縮寫"]].StringValue;
                if (!NameList2.Contains(abbreviation.Trim()))
                {
                    NameList2.Add(abbreviation);
                }
                else
                {
                    NameSb.AppendLine("縮寫重覆:" + abbreviation);
                }
                string hotKey = ws.Cells[x, headers["熱鍵"]].StringValue;
                if (!NameList3.Contains(hotKey.Trim()))
                {
                    NameList3.Add(hotKey);
                }
                else
                {
                    NameSb.AppendLine("熱鍵重覆:" + hotKey);
                }
            }
            if (!string.IsNullOrEmpty(NameSb.ToString()))
            {
                FISCA.Presentation.Controls.MsgBox.Show("匯入缺曠類別發生錯誤:\n" + NameSb.ToString());
                return;
            } 
            #endregion

            for (int x = 1; x <= wb.Worksheets[0].Cells.MaxDataRow; x++) //每一Row
            {
                string name = ws.Cells[x, headers["缺曠名稱"]].StringValue;
                string abbreviation = ws.Cells[x, headers["縮寫"]].StringValue;
                string hotKey = ws.Cells[x, headers["熱鍵"]].StringValue;
                string noabsence = ws.Cells[x, headers["不影響全勤"]].StringValue;

                if (string.IsNullOrEmpty(name.Trim())) //沒有缺曠名稱則跳過
                    continue;

                //if (string.IsNullOrEmpty("" + wb.Worksheets[0].Cells[x, 0].Value))
                //    continue;

                XmlElement absence = doc.CreateElement("Absence");
                root.AppendChild(absence);

                absence.SetAttribute("Name", name.Trim());
                if (!string.IsNullOrEmpty(abbreviation.Trim())) //如果不是空的
                {
                    absence.SetAttribute("Abbreviation", abbreviation.Trim());
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("匯入失敗,縮寫必須有值!!\n此錯誤發生於缺曠名稱:[" + name + "]");
                    return;
                }

                if (CheckHotKey(hotKey.Trim()))
                    absence.SetAttribute("HotKey", hotKey.Trim());
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("匯入失敗,熱鍵必須是英文或數字!\n此錯誤發生於缺曠名稱:[" + name + "]");
                    return;
                }

                absence.SetAttribute("Noabsence", ChangeF(noabsence));
            }

            DSXmlHelper helper = new DSXmlHelper("Lists");
            helper.AddElement("List");
            helper.AddElement("List", "Content", root.OuterXml, true);
            helper.AddElement("List", "Condition");
            helper.AddElement("List/Condition", "Name", Config.LIST_ABSENCE_NAME);

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

            //更新
            try
            {
                Config.Reset(Config.LIST_ABSENCE);
            }
            catch
            {
                FISCA.Presentation.Controls.MsgBox.Show("資料重設失敗!", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            ApplicationLog.Log("學務系統.缺曠類別設定檔", "匯入缺曠類別設定檔", "缺曠類別設定檔，已將缺曠資料管理內容覆蓋匯入。");
            FISCA.Presentation.Controls.MsgBox.Show("匯入成功!\n新設定將於畫面重新開啟時生效!", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            DataGridViewDataInChange = false;
            this.Close();

            #endregion
        }

        private bool CheckHotKey(string hotKey)
        {
            return Pattern.Match(hotKey).Success;
        }

        //private bool CheckInt(string j)
        //{
        //    int HotKeyint;
        //    if (int.TryParse(j, out HotKeyint))
        //    {
        //        return true;
        //    }
        //    else
        //    {
        //        return false;
        //    }
        //}

        private string ChangeF(string u)
        {
            string s = u.Trim().ToUpper();

            string s1 = "True";
            string s2 = "False";

            if (s == "是" || s == "1" || s == "YES" || s == "Y" || s == "影響" || s == "TRUE")
            {
                return s1;
            }
            else if (s == "否" || s == "0" || s == "NO" || s == "N" || s == "不影響" || s == "FALSE")
            {
                return s2;
            }
            else
            {
                return s2;
            }
        }
    }
}
