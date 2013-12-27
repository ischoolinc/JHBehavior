using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using FISCA.Presentation.Controls;
using Campus.Windows;
using K12.Data;
using System.Windows.Forms;
using FISCA.LogAgent;
using System.Drawing;

namespace K12.Behavior.Address
{
    public partial class AddressEditForm : BaseForm
    {
        private StudentData Data;
        private Dictionary<string, int> ColumnIndex = new Dictionary<string, int>();
        private ChangeListener DataListener { get; set; } //DataGridView更新檢查
        private bool DataGridViewDataInChange = false;
        private BackgroundWorker BGW = new BackgroundWorker();
        private Dictionary<string, AddressRecord> dic1 = new Dictionary<string, AddressRecord>(); //Log
        private Dictionary<string, PhoneRecord> dic3 = new Dictionary<string, PhoneRecord>(); //Log

        public AddressEditForm()
        {
            InitializeComponent();
        }

        private void AddressEditForm_Load(object sender, EventArgs e)
        {

            #region Load
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);

            K12.Presentation.NLDPanels.Student.TempSourceChanged += new EventHandler(Student_TempSourceChanged);

            DataListener = new ChangeListener();
            DataListener.Add(new DataGridViewSource(dataGridViewX1));
            DataListener.StatusChanged += new EventHandler<ChangeEventArgs>(DataListener_StatusChanged);

            this.Text = "資料載入中,請稍後...";
            this.Enabled = false;

            BGW.RunWorkerAsync();

            labelX2.Text = "學生待處理：" + K12.Presentation.NLDPanels.Student.TempSource.Count() + "人";
            #endregion
        }

        #region BackgroundWorker
        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            //取得學生資料
            Data = new StudentData();
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Enabled = true;
            this.Text = "聯絡資訊管理";

            comboBoxEx1.SelectedIndex = 0;

            //DataListener.SuspendListen(); //終止變更判斷
            //整理畫面樣式
            //ColumnCheng(); //Column
            //DataCheng(); //Row
            //DataListener.Reset(); //重置內容
            //DataListener.ResumeListen(); //開始判斷 
        }
        #endregion

        #region 資料變更處理
        void DataListener_StatusChanged(object sender, ChangeEventArgs e)
        {
            DataGridViewDataInChange = true;
        }

        private void comboBoxEx1_Enter(object sender, EventArgs e)
        {
            if (DataGridViewDataInChange)
            {
                DialogResult dr = FISCA.Presentation.Controls.MsgBox.Show("您資料已變更,是否要儲存資料?", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button1);
                if (dr == DialogResult.Yes)
                {
                    btnSavePage_Click(null, null);
                }

                DataGridViewDataInChange = false;
            }
        }

        private void comboBoxEx1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataListener.SuspendListen(); //終止變更判斷

            Data.Reset();
            ColumnCheng();
            DataCheng();

            DataListener.Reset();
            DataListener.ResumeListen();
        } 
        #endregion

        private void ColumnCheng()
        {
            #region 調整畫面
            dataGridViewX1.Rows.Clear();
            dataGridViewX1.Columns.Clear();
            ColumnIndex.Clear();

            SetColumnName("ID", 0);
            SetColumnName("班級", 65);
            SetColumnName("座號", 65);
            SetColumnName("姓名", 65);

            if (comboBoxEx1.SelectedIndex == 0)
            {
                #region 地址資料
                SetColumnNameLock("郵遞區號", 90);
                SetColumnNameLock("縣市", 90);
                SetColumnNameLock("鄉鎮", 90);
                SetColumnNameLock("村里", 90);
                SetColumnNameLock("鄰", 90);
                SetColumnNameLock("其他", 250);
                #endregion
            }
            else
            {
                #region 電話資料
                SetColumnNameLock("聯絡電話", 90);
                SetColumnNameLock("其他1", 90);
                SetColumnNameLock("其他2", 90);
                SetColumnNameLock("其他3", 90);
                SetColumnNameLock("手機", 90);
                #endregion
            }
            #endregion
        }

        private void DataCheng()
        {
            #region 填入資料
            dic1.Clear();
            dic3.Clear();
            foreach (StudentRecord each in Data.StudRecordList)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridViewX1);
                row.Tag = false;

                row.Cells[ColumnIndex["ID"]].Value = each.ID;
                row.Cells[ColumnIndex["班級"]].Value = each.Class.Name;
                row.Cells[ColumnIndex["座號"]].Value = each.SeatNo;
                row.Cells[ColumnIndex["姓名"]].Value = each.Name;

                if (comboBoxEx1.SelectedIndex == 0) //地址資料
                {
                    AddressRecord address = Data.Address[each.ID];

                    AddressRecord beforeAddress = new AddressRecord(); //Log
                    row.Cells[ColumnIndex["郵遞區號"]].Value = beforeAddress.Mailing.ZipCode = address.Mailing.ZipCode;
                    row.Cells[ColumnIndex["縣市"]].Value = beforeAddress.Mailing.County = address.Mailing.County;
                    row.Cells[ColumnIndex["鄉鎮"]].Value = beforeAddress.Mailing.Town = address.Mailing.Town;
                    row.Cells[ColumnIndex["村里"]].Value = beforeAddress.Mailing.District = address.Mailing.District;
                    row.Cells[ColumnIndex["鄰"]].Value = beforeAddress.Mailing.Area = address.Mailing.Area;
                    row.Cells[ColumnIndex["其他"]].Value = beforeAddress.Mailing.Detail = address.Mailing.Detail;
                    dic1.Add(each.ID, beforeAddress);//Log

                }
                else //電話資料
                {
                    PhoneRecord phone = Data.Phone[each.ID];

                    PhoneRecord beforePhone = new PhoneRecord(); //Log
                    row.Cells[ColumnIndex["聯絡電話"]].Value = beforePhone.Contact = phone.Contact;
                    row.Cells[ColumnIndex["其他1"]].Value = beforePhone.Phone1 = phone.Phone1;
                    row.Cells[ColumnIndex["其他2"]].Value = beforePhone.Phone2 = phone.Phone2;
                    row.Cells[ColumnIndex["其他3"]].Value = beforePhone.Phone3 = phone.Phone3;
                    row.Cells[ColumnIndex["手機"]].Value = beforePhone.Cell = phone.Cell;
                    dic3.Add(each.ID, beforePhone);//Log
                }

                dataGridViewX1.Rows.Add(row);
            } 
            #endregion
        }

        private void btnSavePage_Click(object sender, EventArgs e)
        {
            btnSavePage.Enabled = false;

            #region 儲存鈕
            if (comboBoxEx1.SelectedIndex == 0)
            {
                #region 地址資料
                Dictionary<string, AddressRecord> dic2 = new Dictionary<string, AddressRecord>();
                foreach (DataGridViewRow each in dataGridViewX1.Rows)
                {
                    bool CHeng = (bool)each.Tag;
                    if (CHeng)
                    {
                        each.Tag = false;
                        if (Data.Address.ContainsKey("" + each.Cells[0].Value))
                        {
                            AddressRecord address = Data.Address["" + each.Cells[ColumnIndex["ID"]].Value];
                            address.Mailing.ZipCode = "" + each.Cells[ColumnIndex["郵遞區號"]].Value;
                            address.Mailing.County = "" + each.Cells[ColumnIndex["縣市"]].Value;
                            address.Mailing.Town = "" + each.Cells[ColumnIndex["鄉鎮"]].Value;
                            address.Mailing.District = "" + each.Cells[ColumnIndex["村里"]].Value;
                            address.Mailing.Area = "" + each.Cells[ColumnIndex["鄰"]].Value;
                            address.Mailing.Detail = "" + each.Cells[ColumnIndex["其他"]].Value;
                            dic2.Add(address.RefStudentID, address); //修改後
                        }
                    }
                }
                try
                {
                    K12.Data.Address.Update(dic2.Values);
                }
                catch (Exception ex)
                {
                    btnSavePage.Enabled = true;
                    FISCA.Presentation.Controls.MsgBox.Show("儲存地址資料,發生錯誤" + ex.Message);
                    return;
                }

                if (dic2.Values.Count != 0)
                {
                    foreach (AddressRecord each in dic2.Values)
                    {
                        StringBuilder sb = new StringBuilder();
                        StudentRecord stud = each.Student;
                        AddressRecord bef = dic1[each.RefStudentID];
                        sb.Append("學生「" + stud.Name + "」");
                        sb.AppendLine("地址資料已被修改。");
                        sb.AppendLine("郵遞區號「" + bef.Mailing.ZipCode + "」改為「" + each.Mailing.ZipCode + "」");
                        sb.AppendLine("縣　　市「" + bef.Mailing.County + "」改為「" + each.Mailing.County + "」");
                        sb.AppendLine("鄉鎮市區「" + bef.Mailing.Town + "」改為「" + each.Mailing.Town + "」");
                        sb.AppendLine("村　　里「" + bef.Mailing.District + "」改為「" + each.Mailing.District + "」");
                        sb.AppendLine("鄰　　　「" + bef.Mailing.Area + "」改為「" + each.Mailing.Area + "」");
                        sb.AppendLine("其　　他「" + bef.Mailing.Detail + "」改為「" + each.Mailing.Detail + "」");
                        ApplicationLog.Log("學務系統.聯絡資訊管理", "修改學生地址資料", "student", stud.ID, sb.ToString());
                    }
                }

                btnSavePage.Enabled = true;
                FISCA.Presentation.Controls.MsgBox.Show("地址資料,儲存成功"); 
                #endregion
            }
            else
            {
                #region 電話資料
                Dictionary<string, PhoneRecord> dic2 = new Dictionary<string, PhoneRecord>();
                foreach (DataGridViewRow each in dataGridViewX1.Rows)
                {
                    bool CHeng = (bool)each.Tag;
                    if (CHeng)
                    {
                        each.Tag = false;
                        if (Data.Address.ContainsKey("" + each.Cells[0].Value))
                        {
                            PhoneRecord phone = Data.Phone["" + each.Cells[ColumnIndex["ID"]].Value];
                            phone.Contact = "" + each.Cells[ColumnIndex["聯絡電話"]].Value;
                            phone.Phone1 = "" + each.Cells[ColumnIndex["其他1"]].Value;
                            phone.Phone2 = "" + each.Cells[ColumnIndex["其他2"]].Value;
                            phone.Phone3 = "" + each.Cells[ColumnIndex["其他3"]].Value;
                            phone.Cell = "" + each.Cells[ColumnIndex["手機"]].Value;
                            dic2.Add(phone.RefStudentID, phone);
                        }
                    }
                }
                try
                {
                    Phone.Update(dic2.Values);
                }
                catch (Exception ex)
                {
                    btnSavePage.Enabled = true;
                    FISCA.Presentation.Controls.MsgBox.Show("儲存電話資料,發生錯誤" + ex.Message);
                    return;
                }

                if (dic2.Values.Count != 0)
                {
                    foreach (PhoneRecord each in dic2.Values)
                    {
                        StringBuilder sb = new StringBuilder();
                        StudentRecord stud = each.Student;
                        PhoneRecord bef = dic3[each.RefStudentID];
                        sb.Append("學生「" + stud.Name + "」");
                        sb.AppendLine("電話資料已被修改。");
                        sb.AppendLine("聯絡電話　「" + bef.Contact + "」改為「" + each.Contact + "」");
                        sb.AppendLine("其他電話一「" + bef.Phone1 + "」改為「" + each.Phone1 + "」");
                        sb.AppendLine("其他電話二「" + bef.Phone2 + "」改為「" + each.Phone2 + "」");
                        sb.AppendLine("其他電話三「" + bef.Phone3 + "」改為「" + each.Phone3 + "」");
                        sb.AppendLine("手　　機　「" + bef.Cell + "」改為「" + each.Cell + "」");
                        ApplicationLog.Log("學務系統.聯絡資訊管理", "修改學生電話資料", "student", stud.ID, sb.ToString());
                    }
                }

                btnSavePage.Enabled = true;
                FISCA.Presentation.Controls.MsgBox.Show("電話資料,儲存成功"); 
                #endregion
            } 
            #endregion

            DataGridViewDataInChange = false;
        }
        
        #region 其他內容
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 新增Column
        /// </summary>
        /// <param name="x"></param>
        private void SetColumnName(string x,int y)
        {
            int NUM = dataGridViewX1.Columns.Add(x, x);
            dataGridViewX1.Columns[NUM].Width = y;
            ColumnIndex.Add(x, NUM);
            SetColumnStyle(NUM);
        }

        /// <summary>
        /// 新增Column,不鎖定
        /// </summary>
        /// <param name="x"></param>
        private void SetColumnNameLock(string x,int y)
        {
            int NUM = dataGridViewX1.Columns.Add(x, x);
            dataGridViewX1.Columns[NUM].Width = y;
            ColumnIndex.Add(x, NUM);
        }

        /// <summary>
        /// 設定預設Column樣式
        /// </summary>
        /// <param name="x"></param>
        private void SetColumnStyle(int x)
        {
            dataGridViewX1.Columns[x].ReadOnly = true;
            dataGridViewX1.Columns[x].DefaultCellStyle.BackColor = Color.LightCyan;

            if (x == 0)
            {
                dataGridViewX1.Columns[x].Visible = false; //如果是ID要隱藏
            }
        }
        #endregion

        private void dataGridViewX1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dataGridViewX1.CurrentRow.Tag = true;
        }

        private void 加入待處理ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            List<string> list = new List<string>();
            foreach (DataGridViewRow row in dataGridViewX1.SelectedRows)
            {
                list.Add("" + row.Cells[0].Value);
            }
            K12.Presentation.NLDPanels.Student.AddToTemp(list);
        }

        private void 清空待處理ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            K12.Presentation.NLDPanels.Student.RemoveFromTemp(K12.Presentation.NLDPanels.Student.TempSource);
        }

        void Student_TempSourceChanged(object sender, EventArgs e)
        {
            labelX2.Text = "學生待處理：" + K12.Presentation.NLDPanels.Student.TempSource.Count() + "人";
        }
    }
}
