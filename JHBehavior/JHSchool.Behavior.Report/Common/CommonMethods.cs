using System;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;
using Aspose.Cells;
using Framework;
using JHSchool;
using JHSchool.Data;
using Campus.ePaperCloud;

namespace JHSchool.Behavior.Report
{
    internal static class CommonMethods
    {
        //Excel報表
        public static void ExcelReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
                return;


            if (e.Error == null)
            {
                string reportName;
                string path;
                Workbook wb;

                object[] result = (object[])e.Result;
                reportName = (string)result[0];
                path = (string)result[1];
                wb = (Workbook)result[2];

                //如果傳入的WorkSheet沒有任何Sheet,就加一個空白的
                if (wb.Worksheets.Count == 0)
                {
                    wb.Worksheets.Add();
                }

                if (File.Exists(path))
                {
                    int i = 1;
                    while (true)
                    {
                        string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                        if (!File.Exists(newPath))
                        {
                            path = newPath;
                            break;
                        }
                    }
                }

                try
                {
                    wb.Save(path, FileFormatType.Excel2003);
                    FISCA.Presentation.MotherForm.SetStatusBarMessage(reportName + "產生完成");
                    System.Diagnostics.Process.Start(path);
                }
                catch (IOException ex1)
                {
                    SaveFileDialog sd = new SaveFileDialog();
                    sd.Title = "另存新檔";
                    sd.FileName = reportName + ".xls";
                    sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
                    if (sd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            wb.Save(sd.FileName, FileFormatType.Excel2003);
                        }
                        catch (IOException ex2)
                        {
                            MsgBox.Show("儲存錯誤,檔案可能開啟中。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        catch (Exception ex2)
                        {
                            SmartSchool.ErrorReporting.ReportingService.ReportException(ex2);
                            MsgBox.Show("列印失敗:\n" + ex1.Message);
                            return;
                        }
                    }
                }
                catch (Exception ex1)
                {
                    SmartSchool.ErrorReporting.ReportingService.ReportException(ex1);
                    MsgBox.Show("列印失敗:\n" + ex1.Message);
                    return;
                }
            }
            else
            {
                MsgBox.Show("列印失敗:\n" + e.Error.Message);
                SmartSchool.ErrorReporting.ReportingService.ReportException(e.Error);
                return;
            }
        }

        //Word報表
        public static void WordReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
                return;

            if (e.Error == null)
            {
                object[] result = (object[])e.Result;

                string reportName = (string)result[0];
                string path = (string)result[1];
                Aspose.Words.Document doc = (Aspose.Words.Document)result[2];
                string path2 = (string)result[3];
                bool PrintStudetnList = (bool)result[4];
                Aspose.Cells.Workbook wb = (Aspose.Cells.Workbook)result[5];

                if (File.Exists(path))
                {
                    int i = 1;
                    while (true)
                    {
                        string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                        if (!File.Exists(newPath))
                        {
                            path = newPath;
                            break;
                        }
                    }
                }

                if (File.Exists(path2))
                {
                    int i = 1;
                    while (true)
                    {
                        string newPath = Path.GetDirectoryName(path2) + "\\" + Path.GetFileNameWithoutExtension(path2) + (i++) + Path.GetExtension(path2);
                        if (!File.Exists(newPath))
                        {
                            path2 = newPath;
                            break;
                        }
                    }
                }

                if (PrintStudetnList)
                {
                    int schoolYear, semester;
                    schoolYear = Convert.ToInt32(School.DefaultSchoolYear);
                    semester = Convert.ToInt32(School.DefaultSemester);
                    MemoryStream memoryStream = new MemoryStream();
                    doc.Save(memoryStream, Aspose.Words.SaveFormat.Doc);
                    ePaperCloud ePaperCloud = new ePaperCloud();
                    ePaperCloud.upload_ePaper(schoolYear, semester, reportName, "", memoryStream, ePaperCloud.ViewerType.Student, ePaperCloud.FormatType.Docx);

                    wb.Save(path2);
                    FISCA.Presentation.MotherForm.SetStatusBarMessage(reportName + "產生完成");
                    System.Diagnostics.Process.Start(path2);
                }
                else
                {
                    int schoolYear, semester;
                    schoolYear = Convert.ToInt32(School.DefaultSchoolYear);
                    semester = Convert.ToInt32(School.DefaultSemester);
                    MemoryStream memoryStream = new MemoryStream();
                    doc.Save(memoryStream, Aspose.Words.SaveFormat.Doc);
                    ePaperCloud ePaperCloud = new ePaperCloud();
                    ePaperCloud.upload_ePaper(schoolYear, semester, reportName, "", memoryStream, ePaperCloud.ViewerType.Student, ePaperCloud.FormatType.Docx);

                    FISCA.Presentation.MotherForm.SetStatusBarMessage(reportName + "產生完成");
                }
            }
            else
            {
                MsgBox.Show("列印失敗:\n" + e.Error.Message);
                SmartSchool.ErrorReporting.ReportingService.ReportException(e.Error);
                return;
            }
        }

        //回報進度
        public static void Report_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("" + e.UserState + "產生中...", e.ProgressPercentage);
        }

        internal static string GetChineseDayOfWeek(DateTime date)
        {
            string dayOfWeek = "";

            switch (date.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    dayOfWeek = "一";
                    break;
                case DayOfWeek.Tuesday:
                    dayOfWeek = "二";
                    break;
                case DayOfWeek.Wednesday:
                    dayOfWeek = "三";
                    break;
                case DayOfWeek.Thursday:
                    dayOfWeek = "四";
                    break;
                case DayOfWeek.Friday:
                    dayOfWeek = "五";
                    break;
                case DayOfWeek.Saturday:
                    dayOfWeek = "六";
                    break;
                case DayOfWeek.Sunday:
                    dayOfWeek = "日";
                    break;
            }

            return dayOfWeek;
        }

        public static int ClassSeatNoComparer(StudentRecord x, StudentRecord y)
        {
            string xx = (x.Class == null ? "" : x.Class.Name) + "::" + x.SeatNo.PadLeft(2, '0');
            string yy = (y.Class == null ? "" : y.Class.Name) + "::" + y.SeatNo.PadLeft(2, '0');

            return xx.CompareTo(yy);
        }

        public static int JHClassSeatNoComparer(JHStudentRecord x, JHStudentRecord y)
        {
            string xx = (x.Class == null ? "" : x.Class.Name) + "::" + (x.SeatNo == null ? "" : x.SeatNo.ToString().PadLeft(2, '0'));
            string yy = (y.Class == null ? "" : y.Class.Name) + "::" + (y.SeatNo == null ? "" : y.SeatNo.ToString().PadLeft(2, '0'));

            return xx.CompareTo(yy);
        }

        public static int ClassComparer(ClassRecord x, ClassRecord y)
        {
            string xx = x.Name;
            string yy = y.Name;

            return xx.CompareTo(yy);
        }

        public static int CusClassComparer(SmartSchool.Customization.Data.ClassRecord x, SmartSchool.Customization.Data.ClassRecord y)
        {
            string xx = x.ClassName;
            string yy = y.ClassName;

            return xx.CompareTo(yy);
        }
    }
}
