using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DevComponents.DotNetBar.Controls;
using System.Windows.Forms;

namespace JHSchool.Behavior.StudentExtendControls
{
    public class CustomDataGridView : DataGridViewX
    {
        #region 此類別CustomDataGridView,繼承DataGridViewX物件,並改寫其內容,將Enter=Tab

        protected override bool ProcessDialogKey(Keys keyData)
        {
            int CurrIndex = this.CurrentCell.ColumnIndex;
            // Extract the key code from the key value. 
            Keys key = (keyData & Keys.KeyCode);

            // Handle the ENTER key as if it were a RIGHT ARROW key. 
            if (key == Keys.Enter)
            {
                foreach (DataGridViewCell cell in this.CurrentRow.Cells)
                {
                    if (cell.ColumnIndex > CurrIndex)
                    {
                        if (cell.ReadOnly)
                        {
                            this.ProcessTabKey(keyData);
                        }
                        else
                        {
                            return this.ProcessTabKey(keyData);
                        }
                    }
                }
            }
            else
            {
                return base.ProcessDialogKey(keyData);
            }
            return this.ProcessTabKey(keyData);
        }

        protected override bool ProcessDataGridViewKey(KeyEventArgs e)
        {
            int CurrIndex = this.CurrentCell.ColumnIndex;


            // Handle the ENTER key as if it were a RIGHT ARROW key. 
            if (e.KeyCode == Keys.Enter)
            {
                foreach (DataGridViewCell cell in this.CurrentRow.Cells)
                {
                    if (cell.ColumnIndex > CurrIndex)
                    {
                        if (cell.ReadOnly)
                        {
                            this.ProcessTabKey(e.KeyData);
                        }
                        else
                        {
                            return this.ProcessTabKey(e.KeyData);
                        }
                    }
                }
            }
            else
            {
                return base.ProcessDataGridViewKey(e);
            }
            return base.ProcessDataGridViewKey(e);
        }

        /// <summary>
        /// 跳一格
        /// </summary>
        public void GoToNEXTCell()
        {
            ProcessDialogKey(Keys.Enter);
        }

        #endregion
    }
}