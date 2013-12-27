using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SmartSchool.Common;
//using SmartSchool.Common;

namespace JHSchool.Behavior.ClassExtendControls.Ribbon
{    
    public partial class DeXingStatistic : FISCA.Presentation.Controls.BaseForm
    {
        private List<string> _classidList;
        private IDeXingExport _exporter;
        public DeXingStatistic(List<string> classidList)
        {
            _classidList = classidList;
            InitializeComponent();
            cboStatisticType.Items.Add("缺曠累計名單");
            cboStatisticType.Items.Add("全勤學生名單");
            cboStatisticType.Items.Add("懲戒特殊表現");
            cboStatisticType.Items.Add("獎勵特殊表現");
            cboStatisticType.SelectedIndex = 0;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cboStatisticType_SelectedIndexChanged(object sender, EventArgs e)
        {            
            switch (cboStatisticType.SelectedItem.ToString())
            {
                case "缺曠累計名單":
                    _exporter = new AttendanceStatistic(_classidList);
                    break;
                case "全勤學生名單":
                    _exporter = new NoAbsenceStatistic(_classidList);
                    break;
                case "懲戒特殊表現":
                    _exporter = new DemeritStatistic(_classidList);
                    break;
                case "獎勵特殊表現":
                    _exporter = new MeritStatistic(_classidList);
                    break;
                default:
                    _exporter = new AttendanceStatistic(_classidList); 
                    break;
            }
            panel1.Controls.Clear();
            Control ctrl = _exporter.MainControl;
            ctrl.Dock = DockStyle.Fill;
            panel1.Controls.Add(ctrl); 
            _exporter.LoadData();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            _exporter.Export();
        }
    }
}