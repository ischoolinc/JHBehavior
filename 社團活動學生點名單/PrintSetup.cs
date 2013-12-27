using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using System.Xml;
using JHSchool;
using System.IO;
using FISCA.Presentation.Controls;
using K12.Data.Configuration;
using FISCA.DSAUtil;

namespace Behavior.ClubActivitiesPointList
{
    public partial class PrintSetup : BaseForm
    {        
        private GetConfigData _Getcd;
        private XmlElement print;
        private string BufferString = null;
        private byte[] _buffer;
        private string base64;

        public PrintSetup(GetConfigData cd)
        {
            InitializeComponent();
            _Getcd = cd;

        }

        private void AttendanceSetup_Load(object sender, EventArgs e)
        {
            if (_Getcd._PrintMode == "true")
            {
                radioButton1.Checked = true;
            }
            else
            {
                radioButton2.Checked = true;
            }

            cbWeekDay.SelectedIndex = _Getcd._Week;

            _buffer = _Getcd._buffer;
            if (_buffer != null)
            {
                base64 = Convert.ToBase64String(_buffer);
            }
            else
            {
                base64 = "";
            }
        }


        private void btnSave_Click(object sender, EventArgs e)
        {
            #region 儲存
            string RedioCheck = "true";
            if (radioButton1.Checked) //true就是預設範本
            {
                RedioCheck = "true";
            }
            else
            {
                RedioCheck = "false";
            }

            DSXmlHelper dsx = new DSXmlHelper("PrintSetup");
            dsx.AddElement("PrintMode");
            dsx.SetAttribute("PrintMode", "bool", RedioCheck);
            dsx.AddElement("PrintTemp");
            dsx.SetAttribute("PrintTemp", "Temp", base64);
            dsx.AddElement("Week");
            dsx.SetAttribute("Week", "day", cbWeekDay.SelectedIndex.ToString()); //預設選星期一

            try
            {
                _Getcd.SavePrint(dsx.BaseElement);
            }
            catch(Exception ex)
            {
                MsgBox.Show("儲存失敗!" + ex.Message);
                return;
            }
            MsgBox.Show("儲存設定成功!");
            this.Close();

            #endregion
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            #region 查看範本
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "(範本)社團活動範本.doc";
            sfd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(Properties.Resources.社團活動範本, 0, Properties.Resources.社團活動範本.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    MsgBox.Show("指定路徑無法存取。", "另存檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            } 
            #endregion
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            #region 查看自定範本
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "(自訂範本)社團活動範本.doc";
            sfd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    if (_buffer == null)
                    {
                        MsgBox.Show("尚無自定範本,請上傳範本!");
                        return;
                    }
                    if (Aspose.Words.Document.DetectFileFormat(new MemoryStream(_buffer)) == Aspose.Words.LoadFormat.Doc)
                        fs.Write(_buffer, 0, _buffer.Length);
                    else
                        fs.Write(Properties.Resources.社團活動範本, 0, Properties.Resources.社團活動範本.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    MsgBox.Show("指定路徑無法存取。", "另存檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            } 
            #endregion
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            #region 上傳自定範本
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "請選擇自訂範本";
            ofd.Filter = "Word檔案 (*.doc)|*.doc";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if (Aspose.Words.Document.DetectFileFormat(ofd.FileName) == Aspose.Words.LoadFormat.Doc)
                    {
                        FileStream fs = new FileStream(ofd.FileName, FileMode.Open);
                        byte[] tempBuffer = new byte[fs.Length];
                        _buffer = tempBuffer;
                        fs.Read(tempBuffer, 0, tempBuffer.Length);
                        base64 = Convert.ToBase64String(tempBuffer);
                        //_isUpload = true;
                        fs.Close();
                        MsgBox.Show("上傳成功。");
                        radioButton2.Checked = true;
                    }
                    else
                        MsgBox.Show("上傳檔案格式不符");
                }
                catch
                {
                    MsgBox.Show("指定路徑無法存取。", "開啟檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            } 
            #endregion
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}