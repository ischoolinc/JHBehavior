using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Aspose.Cells;
using FISCA.Presentation.Controls;

namespace JHSchool.Behavior.Report.班級學生獎懲統計
{
    public partial class MeritDemeritStatistics : BaseForm
    {
        ObjConfig config;

        BackgroundWorker BGW = new BackgroundWorker();

        string SheetName = "Sheet1";

        Dictionary<string, int> ColumnIndexDic = new Dictionary<string, int>();

        int ColumnIndex = 0;

        Workbook wb = new Workbook();

        /// <summary>
        /// 新寫法之班級學生獎懲統計(2/1日)
        /// </summary>
        public MeritDemeritStatistics()
        {
            InitializeComponent();
        }

        private void MeritDemeritStatistics_Load(object sender, EventArgs e)
        {
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);

            dtiStartDate.Value = DateTime.Now.AddDays(-1);
            dtiEndDate.Value = DateTime.Now;

            listView1.Items.Add("大功").Checked = true;
            listView1.Items.Add("小功").Checked = true;
            listView1.Items.Add("嘉獎").Checked = true;
            listView1.Items.Add("大過").Checked = true;
            listView1.Items.Add("小過").Checked = true;
            listView1.Items.Add("警告").Checked = true;
        }

        /// <summary>
        /// 列印
        /// </summary>
        private void btnPrint_Click(object sender, EventArgs e)
        {
            btnPrint.Enabled = false;

            List<string> list = new List<string>();
            foreach (ListViewItem each in listView1.Items)
            {
                if (each.Checked)
                {
                    list.Add(each.Text);
                }
            }

            config = new ObjConfig();
            config.StartDate = dtiStartDate.Value;
            config.EndDate = dtiEndDate.Value;
            config.Cleared = checkBoxX1.Checked;
            config.SelectItems = list;
            config.InsertOrSetup = rbRegisterDate.Checked;

            BGW.RunWorkerAsync();
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            //傳入日期/包含銷過(true)/清單/發生日期(true)
            InfoClass obj = new InfoClass(config);

            e.Result = obj;
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            InfoClass obj = e.Result as InfoClass;

            wb = new Workbook(new MemoryStream(ProjectResource.班級學生獎懲統計),new LoadOptions(LoadFormat.Excel97To2003));

            ColumnIndex = 0;
            ColumnIndexDic.Clear();
            SetColumnName1("班級");
            SetColumnName1("班級人數");
            SetColumnName1("導師姓名");

            #region 獎勵判斷
            if (config.SelectItems.Contains("大功") || config.SelectItems.Contains("小功") || config.SelectItems.Contains("嘉獎"))
            {
                SetColumnName1("獎勵總人數");
            }
            if (config.SelectItems.Contains("大功"))
            {
                SetColumnName2("大功");
                SetColumnName3("大功支數");
                SetColumnName3("大功人次");
            }
            if (config.SelectItems.Contains("小功"))
            {
                SetColumnName2("小功");
                SetColumnName3("小功支數");
                SetColumnName3("小功人次");
            }
            if (config.SelectItems.Contains("嘉獎"))
            {
                SetColumnName2("嘉獎");
                SetColumnName3("嘉獎支數");
                SetColumnName3("嘉獎人次");
            }
            #endregion

            #region 懲戒判斷
            if (config.SelectItems.Contains("大過") || config.SelectItems.Contains("小過") || config.SelectItems.Contains("警告"))
            {
                SetColumnName1("懲戒總人數");
            }
            if (config.SelectItems.Contains("大過"))
            {
                SetColumnName2("大過");
                SetColumnName3("大過支數");
                SetColumnName3("大過人次");
            }
            if (config.SelectItems.Contains("小過"))
            {
                SetColumnName2("小過");
                SetColumnName3("小過支數");
                SetColumnName3("小過人次");
            }
            if (config.SelectItems.Contains("警告"))
            {
                SetColumnName2("警告");
                SetColumnName3("警告支數");
                SetColumnName3("警告人次");
            }
            #endregion

            #region 設定畫面內容
            wb.Worksheets[SheetName].Cells.CreateRange(0, 0, 1, ColumnIndex).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            wb.Worksheets[SheetName].Cells.CreateRange(0, 0, 1, ColumnIndex).SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            wb.Worksheets[SheetName].Cells.CreateRange(0, 0, 1, ColumnIndex).SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            wb.Worksheets[SheetName].Cells.CreateRange(0, 0, 1, ColumnIndex).SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

            wb.Worksheets[SheetName].Cells.CreateRange(1, 0, 1, ColumnIndex).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            wb.Worksheets[SheetName].Cells.CreateRange(1, 0, 1, ColumnIndex).SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            wb.Worksheets[SheetName].Cells.CreateRange(1, 0, 1, ColumnIndex).SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            wb.Worksheets[SheetName].Cells.CreateRange(1, 0, 1, ColumnIndex).SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

            for (int i = 0; i < ColumnIndex - 1; i++)
            {
                wb.Worksheets[SheetName].Cells.CreateRange(0, i, 2, 1).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                wb.Worksheets[SheetName].Cells.CreateRange(0, i, 2, 1).SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                wb.Worksheets[SheetName].Cells.CreateRange(0, i, 2, 1).SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                wb.Worksheets[SheetName].Cells.CreateRange(0, i, 2, 1).SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            }
            #endregion

            int RowIndex = 2;

            #region 迴圈
            foreach (string each in obj.ClassMeritDemeritList.Keys)
            {
                MeritDemeritInfo MDInfo = obj.ClassMeritDemeritList[each];

                wb.Worksheets[SheetName].Cells[RowIndex, ColumnIndexDic["班級"]].PutValue(MDInfo.ClassName);
                wb.Worksheets[SheetName].Cells[RowIndex, ColumnIndexDic["班級人數"]].PutValue(MDInfo.StudentCount);
                wb.Worksheets[SheetName].Cells[RowIndex, ColumnIndexDic["導師姓名"]].PutValue(MDInfo.TeacherName);
                if (config.SelectItems.Contains("大功") || config.SelectItems.Contains("小功") || config.SelectItems.Contains("嘉獎"))
                {
                    wb.Worksheets[SheetName].Cells[RowIndex, ColumnIndexDic["獎勵總人數"]].PutValue(MDInfo.MeritStudentCount);
                }
                if (config.SelectItems.Contains("大功"))
                {
                    wb.Worksheets[SheetName].Cells[RowIndex, ColumnIndexDic["大功支數"]].PutValue(MDInfo.MeritA);
                    wb.Worksheets[SheetName].Cells[RowIndex, ColumnIndexDic["大功人次"]].PutValue(MDInfo.MeritAStudentCount);
                }
                if (config.SelectItems.Contains("小功"))
                {
                    wb.Worksheets[SheetName].Cells[RowIndex, ColumnIndexDic["小功支數"]].PutValue(MDInfo.MeritB);
                    wb.Worksheets[SheetName].Cells[RowIndex, ColumnIndexDic["小功人次"]].PutValue(MDInfo.MeritBStudentCount);
                }
                if (config.SelectItems.Contains("嘉獎"))
                {
                    wb.Worksheets[SheetName].Cells[RowIndex, ColumnIndexDic["嘉獎支數"]].PutValue(MDInfo.MeritC);
                    wb.Worksheets[SheetName].Cells[RowIndex, ColumnIndexDic["嘉獎人次"]].PutValue(MDInfo.MeritCStudentCount);
                }
                if (config.SelectItems.Contains("大過") || config.SelectItems.Contains("小過") || config.SelectItems.Contains("警告"))
                {
                    wb.Worksheets[SheetName].Cells[RowIndex, ColumnIndexDic["懲戒總人數"]].PutValue(MDInfo.DemeritStudentCount);
                }
                if (config.SelectItems.Contains("大過"))
                {
                    wb.Worksheets[SheetName].Cells[RowIndex, ColumnIndexDic["大過支數"]].PutValue(MDInfo.DemeritA);
                    wb.Worksheets[SheetName].Cells[RowIndex, ColumnIndexDic["大過人次"]].PutValue(MDInfo.DemeritAStudentCount);
                }
                if (config.SelectItems.Contains("小過"))
                {
                    wb.Worksheets[SheetName].Cells[RowIndex, ColumnIndexDic["小過支數"]].PutValue(MDInfo.DemeritB);
                    wb.Worksheets[SheetName].Cells[RowIndex, ColumnIndexDic["小過人次"]].PutValue(MDInfo.DemeritBStudentCount);
                }
                if (config.SelectItems.Contains("警告"))
                {
                    wb.Worksheets[SheetName].Cells[RowIndex, ColumnIndexDic["警告支數"]].PutValue(MDInfo.DemeritC);
                    wb.Worksheets[SheetName].Cells[RowIndex, ColumnIndexDic["警告人次"]].PutValue(MDInfo.DemeritCStudentCount);
                }
                RowIndex++;
            }

            wb.Worksheets[SheetName].AutoFitColumns();

            #endregion

            #region 儲存

            btnPrint.Enabled = true;

            SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "班級學生獎懲統計.xlsx";
            sd.Filter = "Excel檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    wb.Save(sd.FileName);
                    System.Diagnostics.Process.Start(sd.FileName);

                }
                catch
                {
                    MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            } 
            #endregion
        }

        private void SetColumnName1(string ColumnName)
        {
            ColumnIndexDic.Add(ColumnName, ColumnIndex);
            wb.Worksheets[SheetName].Cells[1, ColumnIndex].PutValue(ColumnName);

            //wb.Worksheets[SheetName].Cells.Merge(0, ColumnIndex, 2, 1);
            ColumnIndex++;
        }

        private void SetColumnName2(string ColumnName)
        {
            wb.Worksheets[SheetName].Cells[0, ColumnIndex].PutValue(ColumnName);
            wb.Worksheets[SheetName].Cells.Merge(0, ColumnIndex, 1, 2);
        }

        private void SetColumnName3(string ColumnName)
        {
            ColumnIndexDic.Add(ColumnName, ColumnIndex);
            wb.Worksheets[SheetName].Cells[1, ColumnIndex].PutValue(ColumnName);
            ColumnIndex++;
        }

        /// <summary>
        /// 結束畫面
        /// </summary>
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SelectAllChecked_CheckedChanged(object sender, EventArgs e)
        {
            foreach (ListViewItem each in listView1.Items)
            {
                each.Checked = SelectAllChecked.Checked;
            }
        }
    }
}
