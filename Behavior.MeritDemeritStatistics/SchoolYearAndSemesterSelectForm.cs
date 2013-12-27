using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;
using Aspose.Cells;
using System.IO;
using System.Diagnostics;
using FISCA.Presentation;

namespace Behavior.MeritDemeritStatistics
{
    public partial class SchoolYearAndSemesterSelectForm : BaseForm
    {
        BackgroundWorker BGW = new BackgroundWorker();
        int SchoolYear { get; set; }
        int Semester { get; set; }
        bool IsSchoolYear { get; set; }

        //資料收集器
        AcquisitionOfInformation AOI { get; set; }

        public MeritDemeritObj MDObj = new MeritDemeritObj();

        Workbook book;

        public SchoolYearAndSemesterSelectForm()
        {
            InitializeComponent();
        }

        private void SchoolYearAndSemesterSelectForm_Load(object sender, EventArgs e)
        {
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);

            intSchoolYear.Value = int.Parse(School.DefaultSchoolYear);
            intSemester.Value = int.Parse(School.DefaultSemester);
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (!BGW.IsBusy)
            {
                MotherForm.SetStatusBarMessage("全校獎懲人數統計中...");
                btnPrint.Enabled = false;
                SchoolYear = intSchoolYear.Value;
                Semester = intSemester.Value;
                IsSchoolYear = radioButton1.Checked;
                BGW.RunWorkerAsync();
            }
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            //資料收集器
            AOI = new AcquisitionOfInformation(SchoolYear, Semester, IsSchoolYear);

            //資料掃瞄器
            MaterialScanObj MSO = new MaterialScanObj(AOI);

            //資料產生器
            book = MSO.CreateExcel();
            if (IsSchoolYear)
            {
                book.Worksheets[0].Cells[0, 0].PutValue(SchoolYear + "學年度　獎懲人數統計");
            }
            else
            {
                book.Worksheets[0].Cells[0, 0].PutValue(SchoolYear + "學年度　第" + Semester + "學期　獎懲人數統計");
            }
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            RePoint("獎懲人數統計");
            MotherForm.SetStatusBarMessage("獎懲人數統計報表,產生完成!");
            btnPrint.Enabled = true;
            if (AOI.HasNotDividedTheSex.Count > 0)
            {
                DialogResult dr = MsgBox.Show("獎懲人數統計報表,產生完成！\n\n發現未分性別之學生！！\n您是否要將這些學生加入學生待處理\n您可以檢視並瞭解學生資料。", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button1);
                if (dr == DialogResult.Yes)
                {
                    K12.Presentation.NLDPanels.Student.AddToTemp(AOI.HasNotDividedTheSex.Select(X => X.ID).ToList());
                    MsgBox.Show("未分性別之學生已加入學生待處理！");
                }
            }
            else
            {
                MsgBox.Show("獎懲人數統計報表,產生完成！");
            }
        }

        //產生報表/列印報表
        private void RePoint(string Name)
        {
            string path = Path.Combine(Application.StartupPath, "Reports");

            //如果目錄不存在則建立。
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);

            if (IsSchoolYear)
            {
                path = Path.Combine(path, ConvertToValidName(Name + "_" + SchoolYear + "學年度") + ".xls");
            }
            else
            {
                path = Path.Combine(path, ConvertToValidName(Name + "_" + SchoolYear + "學年度第" + Semester + "學期") + ".xls");
            }

            try
            {
                book.Save(path);
            }
            catch (IOException)
            {
                try
                {
                    FileInfo file = new FileInfo(path);
                    string nameTempalte = file.FullName.Replace(file.Extension, "") + "{0}.xls";
                    int count = 1;
                    string fileName = string.Format(nameTempalte, count);
                    while (File.Exists(fileName))
                        fileName = string.Format(nameTempalte, count++);

                    book.Save(fileName);
                    path = fileName;
                }
                catch (Exception ex)
                {
                    MsgBox.Show("檔案儲存失敗:" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch (Exception ex)
            {
                MsgBox.Show("檔案儲存失敗:" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                Process.Start(path);
            }
            catch (Exception ex)
            {
                MsgBox.Show("檔案開啟失敗:" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        //建立檔案時進行判斷並編號
        private string ConvertToValidName(string A1Name)
        {
            char[] invalids = Path.GetInvalidFileNameChars();

            string result = A1Name;
            foreach (char each in invalids)
                result = result.Replace(each, '_');

            return result;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            intSemester.Visible = (!radioButton1.Checked);
            txtSemester.Visible = (!radioButton1.Checked);
        }
    }
}
