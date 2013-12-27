using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using JHSchool;
using JHSchool.Data;
using FISCA.DSAUtil;
using Aspose.Words;
using System.IO;

namespace Behavior.ClubActivitiesPointList
{
    public partial class ClubActivitiesForm : BaseForm
    {
        public const string ConfigName = "社團活動點名單設定檔_Word";

        GetCourseDetail CourseDetail;
        GetConfigData GetCD;
        GetDayList GetDay;
        BackgroundWorker BGW = new BackgroundWorker();

        string Date1;
        string Date2;
        private double FontSize = 10;
        private string FontName = "標楷體";
        private Document _doc;
        private Run _run;

        public ClubActivitiesForm()
        {
            InitializeComponent();
        }

        private void ClubActivitiesForm_Load(object sender, EventArgs e)
        {
            #region Load
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);

            //取得設定狀態
            GetCD = new GetConfigData(ConfigName);
            //初始值
            Date1 = dateTimeInput1.Text = GetCD._Day1;
            Date2 = dateTimeInput2.Text = GetCD._Day2;
            //設定畫面
            ResetDayList(); 
            #endregion
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            #region Start

            if (GetCD._PrintMode == "true")
            {
                if (GetDay.DateTimeList.Count > 10)
                {
                    DialogResult dr = MsgBox.Show("預設之A4紙張樣板,最佳週次為10週次\n是否繼續列印?", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button1);
                    if (dr == DialogResult.No)
                        return;
                }
            }

            btnPrint.Enabled = false;
            BGW.RunWorkerAsync(); 
            #endregion
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            #region 列印

            //儲存使用設定
            DSXmlHelper dsx = new DSXmlHelper("BeforeSetup");
            dsx.AddElement("Day1");
            dsx.SetAttribute("Day1", "day", Date1);
            dsx.AddElement("Day2");
            dsx.SetAttribute("Day2", "day", Date2);
            GetCD.SaveBefore(dsx.BaseElement);

            //取得日期清單
            GetDay = new GetDayList(GetCD._Week.ToString(), Date1, Date2);

            //取得學生課程資料
            CourseDetail = new GetCourseDetail(JHCourse.SelectByIDs(K12.Presentation.NLDPanels.Course.SelectedSource));

            #endregion

            #region 產生Word檔

            //範本檔
            Document template;

            if (GetCD._PrintMode == "true") //true是使用預設
            {
                template = new Document(new MemoryStream(Properties.Resources.社團活動範本));
            }
            else
            {
                if (GetCD._PrintTemp == "")
                {
                    template = new Document(new MemoryStream(Properties.Resources.社團活動範本));
                }
                else
                {
                    template = new Document(new MemoryStream(GetCD._buffer));
                }
            }

            //確認字型
            FontSize = template.Sections[0].Body.Tables[0].Rows[3].Cells[0].Paragraphs[0].Runs[0].Font.Size;
            FontName = template.Sections[0].Body.Tables[0].Rows[3].Cells[0].Paragraphs[0].Runs[0].Font.Name;
            //Cell移動使用
            _run = new Run(template);

            DocumentBuilder DB = new DocumentBuilder(template);
            DB.MoveToMergeField("日期");
            Cell AbsenceCell = DB.CurrentParagraph.ParentNode as Cell; //取得該Cell
            CellSplit(AbsenceCell, GetDay.DateTimeList.Count); //切割器
            int x = 3;
            foreach (DateTime time in GetDay.DateTimeList)
            {
                Write(AbsenceCell, time.ToString("MM/dd"));
                AbsenceCell = GetMoveRightCell(AbsenceCell, 1);
                x++;
            }

            DB.MoveToMergeField("簽名");
            AbsenceCell = DB.CurrentParagraph.ParentNode as Cell; //取得該Cell
            CellSplit(AbsenceCell, GetDay.DateTimeList.Count); //切割器

            //儲存時使用的Document
            _doc = new Document();
            _doc.Sections.Clear();

            foreach (string each in CourseDetail.DicCourseStudent.Keys)
            {
                Document resultdoc = template.Clone(true) as Document;
                _run = new Run(resultdoc);

                #region MailMerge
                List<string> name = new List<string>();
                List<string> value = new List<string>();

                name.Add("學校名稱");
                value.Add(School.ChineseName);

                name.Add("學年度/學期");
                value.Add(School.DefaultSchoolYear + "學年度," + "第" + School.DefaultSemester + "學期");

                name.Add("社團名稱");
                value.Add(CourseDetail.DicCourseByKey[each].Name);

                name.Add("列印日期");
                value.Add(DateTime.Now.ToShortDateString());

                name.Add("列印時間");
                value.Add(DateTime.Now.ToShortTimeString());

                name.Add("人數");
                value.Add(CourseDetail.DicCourseStudent[each].Count.ToString());

                name.Add("資料");
                value.Add(each);

                resultdoc.MailMerge.MergeField += new Aspose.Words.Reporting.MergeFieldEventHandler(MailMerge_MergeField);
                resultdoc.MailMerge.Execute(name.ToArray(), value.ToArray());
                resultdoc.MailMerge.DeleteFields();
                resultdoc.MailMerge.MergeField -= new Aspose.Words.Reporting.MergeFieldEventHandler(MailMerge_MergeField);

                _doc.Sections.Add(_doc.ImportNode(resultdoc.FirstSection, true));

                #endregion
            }


            #endregion
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            #region 當列印完成

            btnPrint.Enabled = true;

            SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "社團活動學生點名單.doc";
            sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    _doc.Save(sd.FileName);
                    System.Diagnostics.Process.Start(sd.FileName);

                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return;
                }
            } 
            #endregion
        }

        void MailMerge_MergeField(object sender, Aspose.Words.Reporting.MergeFieldEventArgs e)
        {
            #region 資料切割
            if (e.FieldName == "資料")
            {
                DocumentBuilder builder = new DocumentBuilder(e.Document);
                builder.MoveToField(e.Field, true);
                e.Field.Remove();

                List<JHStudentRecord> StudList = CourseDetail.DicCourseStudent[e.FieldValue.ToString()];

                Row refrow = builder.CurrentParagraph.ParentNode.ParentNode as Row;
                Cell SplieCell = GetMoveRightCell(refrow.Cells[0], 3);
                CellSplit(SplieCell, GetDay.DateTimeList.Count);

                Row rowtemp = builder.CurrentParagraph.ParentNode.ParentNode.Clone(true) as Row;
                Table table = builder.CurrentParagraph.ParentNode.ParentNode.ParentNode as Table;

                foreach (JHStudentRecord each in StudList)
                {

                    Write(refrow.Cells[0], each.Class != null ? each.Class.Name : "");
                    Write(refrow.Cells[1], each.SeatNo.HasValue ? each.SeatNo.Value.ToString() : "");
                    Write(refrow.Cells[2], each.Name);

                    refrow = table.InsertAfter(rowtemp.Clone(true), refrow) as Row;
                }

                refrow.Remove();
            } 
            #endregion
        }

        private void lkbPrintSetup_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            #region 列印設定
            PrintSetup print = new PrintSetup(GetCD);
            print.ShowDialog();

            GetCD = new GetConfigData(ConfigName);
            ResetDayList();

            #endregion
        }

        #region 其他設定
        private void dateTimeInput1_TextChanged(object sender, EventArgs e)
        {
            //取得日期清單
            if (dateTimeInput1.Text != "" && dateTimeInput2.Text != "")
            {
                ResetDayList();
            }
        }

        private void dateTimeInput2_TextChanged(object sender, EventArgs e)
        {
            //取得日期清單
            if (dateTimeInput1.Text != "" && dateTimeInput2.Text != "")
            {
                ResetDayList();
            }
        }

        private void ResetDayList()
        {
            Date1 = dateTimeInput1.Text;
            Date2 = dateTimeInput2.Text;
            //取得日期清單
            GetDay = new GetDayList(GetCD._Week.ToString(), Date1, Date2);
            lbWeekTime.Text = "社團日：" + CheckWeek(GetCD._Week.ToString()) + "(以上時間區間內,共" + GetDay.DateTimeList.Count.ToString() + "週次)";
        } 
        #endregion

        private string CheckWeek(string x)
        {
            #region 依編號取代為星期
            if (x == "0")
            {
                return "星期一";
            }
            else if (x == "1")
            {
                return "星期二";
            }
            else if (x == "2")
            {
                return "星期三";
            }
            else if (x == "3")
            {
                return "星期四";
            }
            else if (x == "4")
            {
                return "星期五";
            }
            else if (x == "5")
            {
                return "星期六";
            }
            else
            {
                return "星期日";
            }
            #endregion
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Cell切割器
        /// </summary>
        /// <param name="_cell">傳入分割的儲存格</param>
        /// <param name="Count">傳入分割數目</param>
        private void CellSplit(Cell _cell, int Count)
        {
            #region Cell切割器
            double MAXwidth = _cell.CellFormat.Width;
            double Cellwidth = MAXwidth / Count;

            List<Cell> list = new List<Cell>();
            list.Add(_cell);

            Row _row = _cell.ParentNode as Row;
            for (int x = 0; x < Count - 1; x++)
            {
                list.Add((_row.InsertAfter(new Cell(_cell.Document), _cell)) as Cell);
            }

            foreach (Cell each in list)
            {
                each.CellFormat.Width = Cellwidth;
            }
            #endregion
        }

        #region 阿寶友情贊助

        /// <summary>
        /// 以Cell為基準,向下移一格
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        private Cell GetMoveDownCell(Cell cell, int count)
        {
            #region 以Cell為基準,向下移一格
            if (count == 0) return cell;

            Row row = cell.ParentRow;
            int col_index = row.IndexOf(cell);
            Table table = row.ParentTable;
            int row_index = table.Rows.IndexOf(row) + count;

            try
            {
                return table.Rows[row_index].Cells[col_index];
            }
            catch (Exception ex)
            {
                return null;
            }
            #endregion
        }

        /// <summary>
        /// 以Cell為基準,向右移一格
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        private Cell GetMoveRightCell(Cell cell, int count)
        {
            #region 以Cell為基準,向右移一格
            if (count == 0) return cell;

            Row row = cell.ParentRow;
            int col_index = row.IndexOf(cell);
            Table table = row.ParentTable;
            int row_index = table.Rows.IndexOf(row);

            try
            {
                return table.Rows[row_index].Cells[col_index + count];
            }
            catch (Exception ex)
            {
                return null;
            }
            #endregion
        }

        /// <summary>
        /// 以Cell為基準,使用NextSibling向右移一格
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        private Cell GetMoveRightCellByNextSibling(Cell cell, int count)
        {
            #region 以Cell為基準,使用NextSibling向右移一格
            if (count == 0) return cell;

            Node node = cell;
            for (int i = 0; i < count; i++)
                node = node.NextSibling;

            try
            {
                return (Cell)node;
            }
            catch (Exception ex)
            {
                return null;
            }
            #endregion
        }

        /// <summary>
        /// 寫入資料
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="text"></param>
        private void Write(Cell cell, string text)
        {
            #region 寫入資料
            if (cell.FirstParagraph == null)
                cell.Paragraphs.Add(new Paragraph(cell.Document));
            cell.FirstParagraph.Runs.Clear();
            _run.Text = text;
            _run.Font.Size = FontSize;
            _run.Font.Name = FontName;
            cell.FirstParagraph.Runs.Add(_run.Clone(true));
            #endregion
        }

        #endregion
    }
}
