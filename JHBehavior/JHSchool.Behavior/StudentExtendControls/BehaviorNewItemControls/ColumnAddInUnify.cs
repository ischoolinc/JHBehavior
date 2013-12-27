using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace JHSchool.Behavior.StudentExtendControls
{
    class ColumnAddInUnify
    {
        /// <summary>
        /// 傳入Column,會將
        /// </summary>
        public ColumnAddInUnify()
        {

        }

        /// <summary>
        /// 傳入DataGridView,會將Row1加至Row3,當傳入=true為缺曠
        /// </summary>
        /// <param name="DgvRedy"></param>
        public static void SetAddCell(DataGridView DgvRedy,bool Mode)
        {
            DataGridViewRow row1 = DgvRedy.Rows[0];
            DataGridViewRow row2 = DgvRedy.Rows[1];
            DataGridViewRow row3 = DgvRedy.Rows[2];

            foreach (DataGridViewColumn column in DgvRedy.Columns)
            {
                if (column.Index == 0)
                    continue;

                //string Catch = "" + row3.Cells[column.Index].Value;

                int A1 = intTryParse(row1.Cells[column.Index]);
                int B1 = intTryParse(row2.Cells[column.Index]);

                if (Mode) //缺曠引用
                {
                    if (A1 + B1 == 0)
                    {
                        row3.Cells[column.Index].Value = "";
                        continue;
                    }
                }

                int C1 = A1 + B1;

                row3.Cells[column.Index].Value = C1.ToString();
            }
        }


        private static int intTryParse(DataGridViewCell cell)
        {
            int x;
            if (int.TryParse("" + cell.Value, out x))
            {
                return x;
            }
            else
            {
                return 0;
            }
        }
    }
}
