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
using Aspose.Words;
using FISCA.DSAUtil;
using System.Xml;
using System.IO;
using FISCA.Presentation;
using System.Diagnostics;

namespace JHSchool.Behavior.Report
{
    public partial class SpecialRollCallForm : BaseForm
    {
        Dictionary<string, List<JHStudentRecord>> Dic = new Dictionary<string, List<JHStudentRecord>>();

        /// <summary>
        /// 範本設定檔(字串)
        /// </summary>
        private string SpecialRollCallConfig = "JHSchool.Behavior.Report.SpecialRollCall.Config";

        /// <summary>
        /// 傳入學生進行不分班列印
        /// </summary>
        /// <param name="student"></param>
        public SpecialRollCallForm(List<JHStudentRecord> student)
        {
            InitializeComponent();
            student.Sort(ParseStudent);

            Dic.Add("", student);
        }

        /// <summary>
        /// 傳入班級ID進行依班級列印
        /// </summary>
        /// <param name="ClassIDList"></param>
        public SpecialRollCallForm(List<string> ClassIDList)
        {
            InitializeComponent();

            List<JHClassRecord> SortClassList = JHClass.SelectByIDs(ClassIDList);
            SortClassList = SortClassIndex.JHSchoolData_JHClassRecord(SortClassList);

            foreach (JHClassRecord each in SortClassList)
            {
                Dic.Add(each.ID, new List<JHStudentRecord>());
            }

            foreach (JHStudentRecord student in JHStudent.SelectAll())
            {
                if (student.Status == K12.Data.StudentRecord.StudentStatus.一般)
                {
                    if (ClassIDList.Contains(student.RefClassID))
                    {
                        Dic[student.RefClassID].Add(student);
                    }
                }
            }
        }

        private int SortClassName(JHClassRecord x, JHClassRecord y)
        {
            return x.Name.CompareTo(y.Name);
        }

        private void SpecialRollCallForm_Load(object sender, EventArgs e)
        {
            //取得 Period List
            DSResponse dsrsp = JHSchool.Compatibility.Feature.Basic.Config.GetPeriodList();
            foreach (XmlElement var in dsrsp.GetContent().GetElements("Period"))
            {
                listViewEx1.Items.Add(var.GetAttribute("Name")).Checked = true;
            }

            //取得設定檔
            Campus.Report.ReportConfiguration ConfigurationInCadre = new Campus.Report.ReportConfiguration(SpecialRollCallConfig);
            //如果沒有設定過樣板
            if (ConfigurationInCadre.Template == null)
            {
                //預設樣板 & 格式
                ConfigurationInCadre.Template = new Campus.Report.ReportTemplate(ProjectResource.特殊組合點名表, Campus.Report.TemplateType.Word);
                ConfigurationInCadre.Save();
            }

            dateTimeInput2.Value = dateTimeInput1.Value = DateTime.Today;
        }

        Document Template;
        Document PageOne;

        private void buttonX1_Click(object sender, EventArgs e)
        {
            //結束日期不可大於起始日期
            if (dateTimeInput2.Value.CompareTo(dateTimeInput1.Value) == -1)
            {
                MsgBox.Show("起始日期不可大於結束日期!!");
                return;
            }

            TimeSpan ts = dateTimeInput2.Value - dateTimeInput1.Value;

            List<DateTime> TimeList = new List<DateTime>();

            for (int x = 0; x <= ts.Days; x++)
            {
                TimeList.Add(dateTimeInput1.Value.AddDays(x));
            }



            Document doc = new Document();
            doc.Sections.Clear();

            //取得設定檔
            Campus.Report.ReportConfiguration ConfigurationInCadre = new Campus.Report.ReportConfiguration(SpecialRollCallConfig);
            Template = ConfigurationInCadre.Template.ToDocument();

            List<string> list = new List<string>();
            foreach (ListViewItem item in listViewEx1.Items)
            {
                if (item.Checked)
                {
                    list.Add(item.Text);
                }
            }

            if (list.Count == 0)
            {
                MsgBox.Show("請選擇一個節次!!");
                return;
            }

            DocumentBuilder builder = new DocumentBuilder(Template);

            #region 動態節次1
            builder.MoveToMergeField("動態節次1");
            Run delrun1 = new Run(Template);
            delrun1.Font.Name = builder.Font.Name;
            delrun1.Font.Size = builder.Font.Size;

            Cell delcell1 = builder.CurrentParagraph.ParentNode as Cell;
            Row delrow1 = delcell1.ParentRow;
            double totalWidth1 = delcell1.CellFormat.Width;
            delcell1.Remove();
            foreach (string t in list)
            {
                Cell addCell = builder.InsertCell();
                addCell.CellFormat.Width = totalWidth1 / list.Count;
                delrun1.Text = t;
                addCell.FirstParagraph.Runs.Add(delrun1.Clone(true));
                delrow1.Cells.Add(addCell);
            } 
            #endregion

            #region 動態節次2
            builder.MoveToMergeField("動態節次2");
            Run delrun2 = new Run(Template);
            delrun2.Font.Name = builder.Font.Name;
            delrun2.Font.Size = builder.Font.Size;

            Cell delcell2 = builder.CurrentParagraph.ParentNode as Cell;
            Row delrow2 = delcell2.ParentRow;
            double totalWidth2 = delcell2.CellFormat.Width;
            delcell2.Remove();
            foreach (string t in list)
            {
                Cell addCell2 = builder.InsertCell();
                addCell2.CellFormat.Width = totalWidth2 / list.Count;
                addCell2.FirstParagraph.Runs.Add(delrun2.Clone(true));
                delrow2.Cells.Add(addCell2);
            } 
            #endregion

            #region 以班級進行資料列印
            foreach (string ClassID in Dic.Keys)
            {
                foreach (DateTime ti in TimeList)
                {
                    List<JHStudentRecord> valueEach = Dic[ClassID];

                    JHClassRecord cr = JHClass.SelectByID(ClassID);

                    if (valueEach.Count == 0)
                        continue;

                    valueEach.Sort(ParseStudent);
                    PageOne = (Document)builder.Document.Clone(true);

                    _run = new Run(PageOne);

                    DocumentBuilder builder2 = new DocumentBuilder(PageOne);
                    builder2.MoveToMergeField("資料");

                    Table table = builder2.CurrentParagraph.ParentNode.ParentNode.ParentNode as Table;
                    Row _Row = builder2.CurrentParagraph.ParentNode.ParentNode as Row;
                    _Row.Remove();


                    for (int y = 0; y < valueEach.Count; y++)
                    {
                        table.Rows.Insert(3, _Row.Clone(true));
                    }

                    foreach (JHStudentRecord stud in valueEach)
                    {
                        int x = 3;
                        Cell cellText = table.Rows[x].Cells[0];
                        foreach (JHStudentRecord student in valueEach)
                        {
                            Write(cellText, student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "", builder2);
                            cellText = cellText.NextSibling as Cell;
                            Write(cellText, student.StudentNumber, builder2);
                            cellText = cellText.NextSibling as Cell;
                            Write(cellText, student.Name, builder2);
                            cellText = cellText.NextSibling as Cell;
                            Write(cellText, student.Gender, builder2);
                            cellText = cellText.NextSibling as Cell;

                            x++;
                            if (cellText == null)
                                return;

                            if (cellText.ParentRow.NextSibling != null)
                            {
                                cellText = (cellText.ParentRow.NextSibling as Row).Cells[0];
                            }

                        }
                    }

                    List<string> KeyList = new List<string>();
                    KeyList.Add("學校名稱");
                    KeyList.Add("班級");
                    KeyList.Add("年");
                    KeyList.Add("月");
                    KeyList.Add("日");
                    KeyList.Add("星期");

                    List<string> ValueList = new List<string>();
                    ValueList.Add(School.ChineseName);
                    ValueList.Add(cr.Name);

                    if (!string.IsNullOrEmpty(dateTimeInput1.Text))
                    {
                        ValueList.Add("" + ti.Year);
                        ValueList.Add("" + ti.Month);
                        ValueList.Add("" + ti.Day);
                        ValueList.Add(CheckWeek("" + ti.DayOfWeek));
                    }
                    else
                    {
                        ValueList.Add("　　");
                        ValueList.Add("　　");
                        ValueList.Add("　　");
                        ValueList.Add("　　");
                    }

                    PageOne.MailMerge.Execute(KeyList.ToArray(), ValueList.ToArray());
                    doc.Sections.Add(doc.ImportNode(PageOne.FirstSection, true));
                }
            } 
            #endregion

            try
            {
                SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

                SaveFileDialog1.Filter = "Word (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                SaveFileDialog1.FileName = "點名表";

                if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    doc.Save(SaveFileDialog1.FileName);
                    Process.Start(SaveFileDialog1.FileName);
                    MotherForm.SetStatusBarMessage("點名表,列印完成!!");
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("檔案未儲存");
                    return;
                }
            }
            catch
            {
                FISCA.Presentation.Controls.MsgBox.Show("檔案儲存錯誤,請檢查檔案是否開啟中!!");
                MotherForm.SetStatusBarMessage("檔案儲存錯誤,請檢查檔案是否開啟中!!");
            }

        }

        //移動使用
        private Run _run;

        /// <summary>
        /// 寫入資料
        /// </summary>
        private void Write(Cell cell, string text, DocumentBuilder builder)
        {
            if (cell.FirstParagraph == null)
                cell.Paragraphs.Add(new Paragraph(cell.Document));
            cell.FirstParagraph.Runs.Clear();
            _run.Text = text;
            if (builder != null)
            {
                _run.Font.Size = builder.Font.Size;
                _run.Font.Name = builder.Font.Name;
            }
            else
            {
                _run.Font.Size = 10;
                _run.Font.Name = "標楷體";
            }
            cell.FirstParagraph.Runs.Add(_run.Clone(true));
        }

        //排序功能
        private int ParseStudent(JHStudentRecord x, JHStudentRecord y)
        {
            //取得班級名稱
            string Xstring = x.Class != null ? x.Class.Name : "";
            string Ystring = y.Class != null ? y.Class.Name : "";

            //取得座號
            int Xint = x.SeatNo.HasValue ? x.SeatNo.Value : 0;
            int Yint = y.SeatNo.HasValue ? y.SeatNo.Value : 0;
            //班級名稱加:號加座號(靠右對齊補0)
            Xstring += ":" + Xint.ToString().PadLeft(2, '0');
            Ystring += ":" + Yint.ToString().PadLeft(2, '0');

            return Xstring.CompareTo(Ystring);

        }

        private void checkBoxX1_CheckedChanged(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listViewEx1.Items)
            {
                item.Checked = checkBoxX1.Checked;
            }
        }

        private string CheckWeek(string x)
        {
            #region 依編號取代為星期
            if (x == "Monday")
            {
                return "一";
            }
            else if (x == "Tuesday")
            {
                return "二";
            }
            else if (x == "Wednesday")
            {
                return "三";
            }
            else if (x == "Thursday")
            {
                return "四";
            }
            else if (x == "Friday")
            {
                return "五";
            }
            else if (x == "Saturday")
            {
                return "六";
            }
            else
            {
                return "日";
            }
            #endregion
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //取得設定檔
            Campus.Report.ReportConfiguration ConfigurationInCadre = new Campus.Report.ReportConfiguration(SpecialRollCallConfig);
            //畫面內容(範本內容,預設樣式
            Campus.Report.TemplateSettingForm TemplateForm = new Campus.Report.TemplateSettingForm(ConfigurationInCadre.Template, new Campus.Report.ReportTemplate(ProjectResource.特殊組合點名表, Campus.Report.TemplateType.Word));
            //預設名稱
            TemplateForm.DefaultFileName = "點名表(Word範本)";
            //如果回傳為OK
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                //設定後樣試,回傳
                ConfigurationInCadre.Template = TemplateForm.Template;
                //儲存
                ConfigurationInCadre.Save();
            }
        }

        private void dateTimeInput1_TextChanged(object sender, EventArgs e)
        {
            TimeSpan ts = dateTimeInput2.Value - dateTimeInput1.Value;
            if (ts.Days >= 0)
            {
                groupPanel1.Text = "列印天數：" + (ts.Days + 1).ToString() + "(每班張數)";
            }
            else
            {
                groupPanel1.Text = "列印天數：無法列印";
            }
        }

        private void dateTimeInput2_TextChanged(object sender, EventArgs e)
        {
            TimeSpan ts = dateTimeInput2.Value - dateTimeInput1.Value;
            if (ts.Days >= 0)
            {
                groupPanel1.Text = "列印天數：" + (ts.Days + 1).ToString() + "(每班張數)";
            }
            else
            {
                groupPanel1.Text = "列印天數：無法列印";
            }
        }
    }
}
