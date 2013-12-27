using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using Aspose.Cells;
using FISCA.DSAUtil;
using Framework;

namespace JHSchool.Behavior.Report.班級點名表
{
    internal class Report : IReport
    {
        #region IReport 成員

        private BackgroundWorker _BGWClassStudentAttendance;

        public void Print()
        {
            if (Class.Instance.SelectedList.Count == 0)
                return;

            List<string> config = new List<string>();

            ClassStudentAttendanceSelectPeriodForm form = new ClassStudentAttendanceSelectPeriodForm("班級點名表_節次設定");

            if (form.ShowDialog() == DialogResult.OK)
            {
                XmlElement preferenceData = User.Configuration["班級點名表_節次設定"].GetXml("XmlData", null);

                if (preferenceData == null)
                {
                    MsgBox.Show("第一次使用班級點名表請先設定節次。");
                    return;
                }
                else
                {
                    foreach (XmlElement period in preferenceData.SelectNodes("Period"))
                    {
                        string name = period.GetAttribute("Name");
                        config.Add(name);
                    }
                }

                FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化班級點名表...");

                object[] args = new object[] { config };

                _BGWClassStudentAttendance = new BackgroundWorker();
                _BGWClassStudentAttendance.DoWork += new DoWorkEventHandler(_BGWClassStudentAttendance_DoWork);
                _BGWClassStudentAttendance.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CommonMethods.ExcelReport_RunWorkerCompleted);
                _BGWClassStudentAttendance.ProgressChanged += new ProgressChangedEventHandler(CommonMethods.Report_ProgressChanged);
                _BGWClassStudentAttendance.WorkerReportsProgress = true;
                _BGWClassStudentAttendance.RunWorkerAsync(args);
            }
        }

        void _BGWClassStudentAttendance_DoWork(object sender, DoWorkEventArgs e)
        {
            string reportName = "班級點名表";

            object[] args = e.Argument as object[];

            List<string> config = args[0] as List<string>;

            #region 快取資訊

            List<ClassRecord> selectedClass = Class.Instance.SelectedList;
            selectedClass = SortClassIndex.JHSchool_ClassRecord(selectedClass);

            Dictionary<string, List<StudentRecord>> classStudentList = new Dictionary<string, List<StudentRecord>>();

            //學生人數
            int currentStudentCount = 1;
            int totalStudentNumber = 0;

            //紀錄每一個 Column 的 Index
            Dictionary<string, int> columnTable = new Dictionary<string, int>();

            //使用者定義的節次列表
            List<string> userDefinedPeriodList = new List<string>();

            //節次對照表
            List<string> periodList = new List<string>();

            //依據 ClassID 建立班級學生清單
            foreach (ClassRecord aClass in selectedClass)
            {
                List<StudentRecord> classStudent = aClass.Students.GetInSchoolStudents();
                if (!classStudentList.ContainsKey(aClass.ID))
                    classStudentList.Add(aClass.ID, classStudent);
                totalStudentNumber += classStudent.Count;
            }

            //取得 Period List
            DSResponse dsrsp = JHSchool.Compatibility.Feature.Basic.Config.GetPeriodList();
            foreach (XmlElement var in dsrsp.GetContent().GetElements("Period"))
            {
                if (!periodList.Contains(var.GetAttribute("Name")))
                    periodList.Add(var.GetAttribute("Name"));
            }

            //套用使用者的設定
            userDefinedPeriodList = config;

            #endregion

            #region 產生範本

            Workbook template = new Workbook();
            template.Open(new MemoryStream(ProjectResource.班級點名表), FileFormatType.Excel2003);

            Range tempStudent = template.Worksheets[0].Cells.CreateRange(0, 4, true);
            Range tempEachColumn = template.Worksheets[0].Cells.CreateRange(4, 1, true);

            Workbook prototype = new Workbook();
            prototype.Copy(template);

            prototype.Worksheets[0].Cells.CreateRange(0, 4, true).Copy(tempStudent);

            int titleRow = 2;
            int colIndex = 4;

            int startIndex = colIndex;
            int endIndex;
            int columnNumber;

            //根據使用者定義的節次動態產生欄位
            foreach (string period in userDefinedPeriodList)
            {
                prototype.Worksheets[0].Cells.CreateRange(colIndex, 1, true).Copy(tempEachColumn);
                prototype.Worksheets[0].Cells[titleRow, colIndex].PutValue(period);
                columnTable.Add(period, colIndex - 4);
                colIndex++;
            }

            endIndex = colIndex;
            columnNumber = endIndex - startIndex;
            if (columnNumber == 0)
                columnNumber = 1;

            prototype.Worksheets[0].Cells.CreateRange(0, 0, 1, endIndex).Merge();
            if (endIndex - 3 > 0)
                prototype.Worksheets[0].Cells.CreateRange(1, 3, 1, endIndex - 3).Merge();

            Range prototypeRow = prototype.Worksheets[0].Cells.CreateRange(3, 1, false);
            Range prototypeHeader = prototype.Worksheets[0].Cells.CreateRange(0, 3, false);

            #endregion

            #region 產生報表

            Workbook wb = new Workbook();
            wb.Copy(prototype);

            int index = 0;
            int dataIndex = 0;

            foreach (ClassRecord classInfo in selectedClass)
            {
                List<StudentRecord> classStudent = classStudentList[classInfo.ID];

                //如果不是第一頁，就在上一頁的資料列下邊加黑線
                if (index != 0)
                    wb.Worksheets[0].Cells.CreateRange(index - 1, 0, 1, endIndex).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);

                //複製 Header
                wb.Worksheets[0].Cells.CreateRange(index, 5, false).Copy(prototypeHeader);

                //填寫基本資料
                wb.Worksheets[0].Cells[index, 0].PutValue(School.ChineseName + " 班級點名表");
                wb.Worksheets[0].Cells[index + 1, 0].PutValue("班級：" + classInfo.Name);
                wb.Worksheets[0].Cells[index + 1, 3].PutValue("年　　　月　　　日　星期：　　　");

                dataIndex = index + 3;

                int studentCount = 0;
                while (studentCount < classStudent.Count)
                {
                    //複製每一個 row
                    wb.Worksheets[0].Cells.CreateRange(dataIndex, 1, false).Copy(prototypeRow);
                    //if (studentCount % 5 == 0 && studentCount != 0)
                    //{
                    //    Range eachFiveRow = wb.Worksheets[0].Cells.CreateRange(dataIndex, 0, 1, dayStartIndex);
                    //    eachFiveRow.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Double, Color.Black);
                    //}

                    //填寫學生資料
                    StudentRecord student = classStudent[studentCount];
                    wb.Worksheets[0].Cells[dataIndex, 0].PutValue(student.SeatNo);
                    wb.Worksheets[0].Cells[dataIndex, 1].PutValue(student.StudentNumber);
                    wb.Worksheets[0].Cells[dataIndex, 2].PutValue(student.Name);
                    wb.Worksheets[0].Cells[dataIndex, 3].PutValue(student.Gender);

                    studentCount++;
                    dataIndex++;
                    _BGWClassStudentAttendance.ReportProgress((int)(((double)currentStudentCount++ * 100.0) / (double)totalStudentNumber));
                }

                //資料列上邊加上黑線
                wb.Worksheets[0].Cells.CreateRange(index + 2, 0, 1, endIndex).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);

                //表格最右邊加上黑線
                wb.Worksheets[0].Cells.CreateRange(index + 2, endIndex - 1, studentCount + 1, 1).SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);

                index = dataIndex;

                //設定分頁
                wb.Worksheets[0].HPageBreaks.Add(index, endIndex);
            }

            //最後一頁的資料列下邊加上黑線
            wb.Worksheets[0].Cells.CreateRange(dataIndex - 1, 0, 1, endIndex).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);

            #endregion

            string path = Path.Combine(Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".xlt");
            e.Result = new object[] { reportName, path, wb };
        }

        #endregion
    }
}
