using Aspose.Cells;
using FISCA.Presentation.Controls;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Transfer_Excel
{
    internal class CommonMethods
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

            string xx = x.Class == null ? "" : x.Class.Name + "::";
            xx += x.SeatNo.HasValue ? x.SeatNo.Value.ToString().PadLeft(2, '0') : "00";

            string yy = (y.Class == null ? "" : y.Class.Name) + "::";
            yy += y.SeatNo.HasValue ? y.SeatNo.Value.ToString().PadLeft(2, '0') : "00";
            return xx.CompareTo(yy);
        }

        public static int ClassComparer(ClassRecord x, ClassRecord y)
        {
            string xx = x.Name;
            string yy = y.Name;

            return xx.CompareTo(yy);
        }
    }

}
