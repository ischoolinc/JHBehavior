using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JHSchool.Editor;
using Framework;

namespace JHSchool.Behavior.Editor
{
    public class ClearDemeritRecordEditor : DemeritRecordEditor
    {
        public ClearDemeritRecordEditor(DemeritRecord demeritRecord):base(demeritRecord)
        {
            ClearDate = demeritRecord.ClearDate;       //銷過日期
            ClearReason = demeritRecord.ClearReason;   //銷過事由
            Cleared = demeritRecord.Cleared;           //銷過
        }

        ///// <summary>
        ///// 銷過日期
        ///// </summary>
        //public string ClearDate { get; set; }
        ///// <summary>
        ///// 銷過事由
        ///// </summary>
        //public string ClearReason { get; set; }
        ///// <summary>
        ///// 是否銷過
        ///// </summary>
        //public string Cleared { get; set; }

        public override void Save()
        {
            if (this.EditorStatus == EditorStatus.Update)
                Feature.EditDemerit.SaveClearDemeritRecordEditor(this);
        }

    }

    public static class ClearDemeritRecordEditor_ExtendFunctions
    {
        public static ClearDemeritRecordEditor GetClearDemeritEditor(this DemeritRecord demeritRecord)
        {
            return new ClearDemeritRecordEditor(demeritRecord);
        }

        public static void SaveAll(this IEnumerable<ClearDemeritRecordEditor> editors)
        {
            Feature.EditDemerit.SaveClearDemeritRecordEditors(editors);
        }
    }
}
