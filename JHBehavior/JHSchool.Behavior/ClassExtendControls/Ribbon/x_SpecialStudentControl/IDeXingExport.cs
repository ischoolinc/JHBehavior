using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace JHSchool.Behavior.ClassExtendControls.Ribbon
{
    interface IDeXingExport
    {
        Control MainControl { get;}
        void LoadData();
        void Export();
    }
}
