using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Transfer_Excel
{
    public static class tool
    {
        public static Range CopyStyle(Range cell, Range range)
        {
            cell.Copy(range);
            cell.CopyStyle(range);
            cell.CopyData(range);
            return cell;
        }
    }
}
