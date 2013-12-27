using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Framework;
using FISCA.DSAUtil;
using System.Xml;
using FISCA.LogAgent;

namespace JHSchool.Behavior.StuAdminExtendControls.Ribbon
{
    public partial class PerformanceDegreeForm : FISCA.Presentation.Controls.BaseForm
    {
        //    <DailyBehavior Name="日常行為表現">
        //        <Item Name="愛整潔" Index="....."/>
        //        <Item Name="其他1" Index="....."/>
        //        <Item Name="其他2" Index="....."/>
        //        <Item Name="其他3" Index="....."/>
        //        <PerformanceDegree>
        //            <Mapping Degree="4" Desc="完全符合"/>
        //            <Mapping Degree="3" Desc="大部份符合"/>
        //            <Mapping Degree="2" Desc="部份符合"/>
        //        </PerformanceDegree>
        //    </DailyBehavior>

        public PerformanceDegreeForm()
        {
            //日常行為表現,表現程度
            InitializeComponent();

            ConfigData cd = School.Configuration["DLBehaviorConfig"];

            if (cd.Contains("DailyBehavior"))
            {
                XmlElement dailyBehavior = XmlHelper.LoadXml(cd["DailyBehavior"]);

                foreach (XmlElement mapping in dailyBehavior.SelectNodes("PerformanceDegree/Mapping"))
                {
                    dgvDailyBehaviorMapping.Rows.Add(mapping.GetAttribute("Degree"), mapping.GetAttribute("Desc"));
                }

                cd["DailyBehavior"] = dailyBehavior.OuterXml;
            }

            cd.Save();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (DataCheck())
            {
                MsgBox.Show("資料有誤,請修正後再儲存!!");
                return;
            }


            //取得資料
            ConfigData cd = School.Configuration["DLBehaviorConfig"];

            XmlElement dailyBehavior;

            if (cd.Contains("DailyBehavior"))
            {
                //取得替換的部份
                dailyBehavior = XmlHelper.LoadXml(cd["DailyBehavior"]);
            }
            else
            {
                DSXmlHelper DSXH = new DSXmlHelper("DailyBehavior");
                DSXH.AddElement("PerformanceDegree");
                dailyBehavior = DSXH.BaseElement;
            }

            //取得
            XmlElement db = dailyBehavior.SelectSingleNode("PerformanceDegree") as XmlElement;

            //新增一個資料內容
            DSXmlHelper NewXml = new DSXmlHelper("PerformanceDegree");

            //建立內容
            foreach (DataGridViewRow row in dgvDailyBehaviorMapping.Rows)
            {
                if (row.IsNewRow)
                    continue;

                XmlElement m = NewXml.AddElement("Mapping");
                m.SetAttribute("Degree", "" + row.Cells[0].Value);
                m.SetAttribute("Desc", "" + row.Cells[1].Value);
            }

            //組合內容的關連動作
            XmlNode Xn = dailyBehavior.OwnerDocument.ImportNode(NewXml.BaseElement, true);

            //取代
            dailyBehavior.ReplaceChild(Xn, db);

            //更新替換的部份
            cd["DailyBehavior"] = dailyBehavior.OuterXml;

            try
            {
                //儲存
                cd.Save();
                School.Configuration.Sync("DLBehaviorConfig"); //重置設定檔
            }
            catch (Exception ex)
            {
                MsgBox.Show("儲存資料失敗" + ex.Message);
                return;
            }
            ApplicationLog.Log("學務系統.表現程度對照表", "修改表現程度對照表", "「表現程度對照表」已被修改。");
            this.Close();
        }

        private bool DataCheck()
        {
            foreach (DataGridViewRow row in dgvDailyBehaviorMapping.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.ErrorText))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //檢查機制
        private void dgvDailyBehaviorMapping_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewCell DailyCell = dgvDailyBehaviorMapping.Rows[e.RowIndex].Cells[e.ColumnIndex];

            CheckDataIsNull(DailyCell);

            CheckIs重覆();
        }

        private void CheckIs重覆()
        {
            List<string> list = new List<string>();
            foreach (DataGridViewRow row in dgvDailyBehaviorMapping.Rows)
            {
                if (row.IsNewRow)
                    continue;

                if (!list.Contains("" + row.Cells[0].Value))
                {
                    list.Add("" + row.Cells[0].Value);
                }
                else
                {
                    row.Cells[0].ErrorText += "資料重覆!!　";
                }
            }
        }

        private void CheckDataIsNull(DataGridViewCell DailyCell)
        {
            if (string.IsNullOrEmpty("" + DailyCell.Value))
            {
                DailyCell.ErrorText += "必須輸入資料　";
            }
            else
            {
                DailyCell.ErrorText = "";
            }
        }
    }
}
