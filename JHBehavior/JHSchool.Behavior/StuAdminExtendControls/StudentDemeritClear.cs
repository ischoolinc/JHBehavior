using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using JHSchool.Data;
using Aspose.Cells;
using System.IO;
using System.Diagnostics;

namespace JHSchool.Behavior.StuAdminExtendControls
{
    public partial class StudentDemeritClear : BaseForm
    {
        Workbook book;

        public StudentDemeritClear()
        {
            InitializeComponent();
        }

        private void StudentDemeritClear_Load(object sender, EventArgs e)
        {
            dateTimeInput1.Text = DateTime.Now.AddDays(-6).ToShortDateString();
            dateTimeInput2.Text = DateTime.Now.ToShortDateString();

            buttonX1.Focus();
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            List<JHDemeritRecord> Clearlist = new List<JHDemeritRecord>();
            foreach (JHDemeritRecord each in JHDemerit.SelectAll())
            {
                if (each.Cleared == "是") //銷過
                {
                    if (each.ClearDate.HasValue) //有銷過日期
                    {
                        int CompareStartTime = each.ClearDate.Value.CompareTo(dateTimeInput1.Value); //比開始時間大於/等於(0 or 1)
                        int CompareEndTime = each.ClearDate.Value.CompareTo(dateTimeInput2.Value); //比結束時間小於/等於(0 or -1)

                        if (CompareStartTime != -1 && CompareEndTime != 1)
                        {
                            Clearlist.Add(each);
                        }
                    }
                }
            }

            book = new Workbook();
            book.Worksheets.Clear();
            int SHEETIndex = book.Worksheets.Add();
            Worksheet sheet2 = book.Worksheets[SHEETIndex];
            sheet2.Name = "銷過記錄清單";

            Cell titleCell = sheet2.Cells["A1"];
            titleCell.Style.Borders.SetColor(Color.Black);
            titleCell.PutValue(School.ChineseName + "　銷過記錄清單");
            titleCell.Style.HorizontalAlignment = TextAlignmentType.Center;
            sheet2.Cells.Merge(0, 0, 1, 15);

            FormatCell(sheet2.Cells["A2"], "班級");
            FormatCell(sheet2.Cells["B2"], "座號");
            FormatCell(sheet2.Cells["C2"], "姓名");
            FormatCell(sheet2.Cells["D2"], "學號");
            FormatCell(sheet2.Cells["E2"], "學年度");
            FormatCell(sheet2.Cells["F2"], "學期");
            FormatCell(sheet2.Cells["G2"], "發生日期");
            FormatCell(sheet2.Cells["H2"], "大過");
            FormatCell(sheet2.Cells["I2"], "小過");
            FormatCell(sheet2.Cells["J2"], "警告");
            FormatCell(sheet2.Cells["K2"], "事由");
            FormatCell(sheet2.Cells["L2"], "是否銷過");
            FormatCell(sheet2.Cells["M2"], "銷過日期");
            FormatCell(sheet2.Cells["N2"], "銷過事由");
            FormatCell(sheet2.Cells["O2"], "登錄日期");

            Clearlist.Sort(new SortClass().SortDemeritRecord);

            int ri = 3;
            foreach (JHDemeritRecord each in Clearlist)
            {
                JHStudentRecord student = JHStudent.SelectByID(each.RefStudentID); //取得學生

                if (student.Status != K12.Data.StudentRecord.StudentStatus.一般)
                    continue;

                FormatCell(sheet2.Cells["A" + ri], student.Class != null ? student.Class.Name : "");
                FormatCell(sheet2.Cells["B" + ri], student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "");
                FormatCell(sheet2.Cells["C" + ri], student.Name);
                FormatCell(sheet2.Cells["D" + ri], student.StudentNumber);
                FormatCell(sheet2.Cells["E" + ri], each.SchoolYear.ToString());
                FormatCell(sheet2.Cells["F" + ri], each.Semester.ToString());
                FormatCell(sheet2.Cells["G" + ri], each.OccurDate.ToShortDateString());
                FormatCell(sheet2.Cells["H" + ri], each.DemeritA.HasValue ? each.DemeritA.Value.ToString() : "");
                FormatCell(sheet2.Cells["I" + ri], each.DemeritA.HasValue ? each.DemeritB.Value.ToString() : "");
                FormatCell(sheet2.Cells["J" + ri], each.DemeritA.HasValue ? each.DemeritC.Value.ToString() : "");
                FormatCell(sheet2.Cells["K" + ri], each.Reason);
                FormatCell(sheet2.Cells["L" + ri], each.Cleared);
                FormatCell(sheet2.Cells["M" + ri], each.ClearDate.HasValue ? each.ClearDate.Value.ToShortDateString() : "");
                FormatCell(sheet2.Cells["N" + ri], each.ClearReason);
                FormatCell(sheet2.Cells["O" + ri], each.RegisterDate.HasValue ? each.RegisterDate.Value.ToShortDateString() : "");

                ri++;
            }

            sheet2.AutoFitColumns();

            RePoint("銷過記錄清單");
        }

        private int SortStudent(JHDemeritRecord x, JHDemeritRecord y)
        {
            JHStudentRecord student1 = x.Student;
            JHStudentRecord student2 = y.Student;

            string ClassName1 = student1.Class != null ? student1.Class.Name : "";
            ClassName1 = ClassName1.PadLeft(5, '0');
            string ClassName2 = student2.Class != null ? student2.Class.Name : "";
            ClassName2 = ClassName2.PadLeft(5, '0');

            string Sean1 = student1.SeatNo.HasValue ? student1.SeatNo.Value.ToString() : "";
            Sean1 = Sean1.PadLeft(3, '0');
            string Sean2 = student2.SeatNo.HasValue ? student2.SeatNo.Value.ToString() : "";
            Sean2 = Sean2.PadLeft(3, '0');

            ClassName1 += Sean1;
            ClassName2 += Sean2;

            return ClassName1.CompareTo(ClassName2);
        }

        //產生報表/列印報表
        private void RePoint(string Name)
        {

            foreach (Worksheet sheet in book.Worksheets)
            {
                sheet.AutoFitColumns();
            }

            string path = Path.Combine(Application.StartupPath, "Reports");

            //如果目錄不存在則建立。
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);

            path = Path.Combine(path, ConvertToValidName("銷過記錄清單") + ".xls");
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
                    FISCA.Presentation.Controls.MsgBox.Show("檔案儲存失敗:" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("檔案儲存失敗:" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                Process.Start(path);
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("檔案開啟失敗:" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        //格式化Cell
        private void FormatCell(Cell cell, string value)
        {
            cell.PutValue(value);
            cell.Style.Borders.SetStyle(CellBorderType.Hair);
            cell.Style.Borders.SetColor(Color.Black);
            cell.Style.Borders.DiagonalStyle = CellBorderType.None;
            cell.Style.HorizontalAlignment = TextAlignmentType.Center;
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
