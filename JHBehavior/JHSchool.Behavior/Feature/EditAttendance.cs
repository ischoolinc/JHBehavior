using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Framework;
using FISCA.DSAUtil;
using JHSchool.Editor;
using JHSchool.Behavior.Editor;
using FISCA.Authentication;

namespace JHSchool.Behavior.Feature
{
    /// <summary>
    /// 編輯學生缺曠紀錄
    /// </summary>
    public static class EditAttendance
    {
        private const string INSERT_ATTENDANCE = "SmartSchool.Student.Attendance.Insert";
        private const string UPDATE_ATTENDANCE = "SmartSchool.Student.Attendance.Update";
        private const string DELETE_ATTENDANCE = "SmartSchool.Student.Attendance.Delete";

        /// <summary>
        /// 批次儲存多個Editor
        /// </summary>
        /// <param name="editors"></param>
        internal static void SaveAttendanceRecordEditor(IEnumerable<AttendanceRecordEditor> editors)
        {
            MultiThreadWorker<AttendanceRecordEditor> worker = new MultiThreadWorker<AttendanceRecordEditor>();
            worker.MaxThreads = 3;
            worker.PackageSize = 100;

            worker.PackageWorker += delegate(object sender, PackageWorkEventArgs<AttendanceRecordEditor> e)
            {
                List<string> primarykeys = new List<string>();

                DSXmlHelper insertHelper = new DSXmlHelper("Request");
                DSXmlHelper updateHelper = new DSXmlHelper("Request");
                DSXmlHelper deleteHelper = new DSXmlHelper("Request");

                bool hasInsert = false;
                bool hasUpdate = false;
                bool hasRemove = false;

                foreach (var editor in e.List)
                {
                     primarykeys.Add(editor.RefStudentID);

                    switch (editor.EditorStatus)
                    {
                        #region 新增
                        case EditorStatus.Insert:
                            insertHelper.AddElement("Attendance");
                            insertHelper.AddElement("Attendance", "Field");
                            insertHelper.AddElement("Attendance/Field", "RefStudentID", editor.RefStudentID);
                            insertHelper.AddElement("Attendance/Field", "SchoolYear", editor.SchoolYear);
                            insertHelper.AddElement("Attendance/Field", "Semester", editor.Semester);
                            insertHelper.AddElement("Attendance/Field", "OccurDate", editor.OccurDate);

                            insertHelper.AddElement("Attendance/Field", "Detail");
                            insertHelper.AddElement("Attendance/Field/Detail", "Attendance");
                            foreach (var period in editor.PeriodDetail) 
                            {
                                XmlElement node = insertHelper.AddElement("Attendance/Field/Detail/Attendance", "Period", period.Period);
                                node.SetAttribute("AbsenceType", period.AbsenceType);
                                //node.SetAttribute("AttendanceType", period.AttendanceType);
                            }

                            hasInsert = true;
                            break;
                        #endregion
                        #region 修改
                        case EditorStatus.Update:
                            updateHelper.AddElement("Attendance");
                            updateHelper.AddElement("Attendance", "Field");
                            updateHelper.AddElement("Attendance/Field", "RefStudentID", editor.RefStudentID);
                            updateHelper.AddElement("Attendance/Field", "SchoolYear", editor.SchoolYear);
                            updateHelper.AddElement("Attendance/Field", "Semester", editor.Semester);
                            updateHelper.AddElement("Attendance/Field", "OccurDate", editor.OccurDate);

                            updateHelper.AddElement("Attendance/Field", "Detail");
                            updateHelper.AddElement("Attendance/Field/Detail", "Attendance");
                            foreach (var period in editor.PeriodDetail)
                            {
                                XmlElement node = updateHelper.AddElement("Attendance/Field/Detail/Attendance", "Period", period.Period);
                                node.SetAttribute("AbsenceType", period.AbsenceType);
                                //node.SetAttribute("AttendanceType", period.AttendanceType);
                            }

                            updateHelper.AddElement("Attendance", "Condition");
                            updateHelper.AddElement("Attendance/Condition", "ID", editor.ID);
                            updateHelper.AddElement("Attendance/Condition", "RefStudentID", editor.RefStudentID);

                            hasUpdate = true;
                            break;
                        #endregion
                        #region 刪除
                        case EditorStatus.Delete:
                            deleteHelper.AddElement("Attendance");
                            deleteHelper.AddElement("Attendance", "ID", editor.ID);

                            hasRemove = true;
                            break;
                        #endregion
                    }
                }

                if (hasInsert)
                    DSAServices.CallService(INSERT_ATTENDANCE, new DSRequest(insertHelper.BaseElement));

                if (hasUpdate)
                    DSAServices.CallService(UPDATE_ATTENDANCE, new DSRequest(updateHelper.BaseElement));

                if (hasRemove)
                    DSAServices.CallService(DELETE_ATTENDANCE, new DSRequest(deleteHelper.BaseElement));

                if (primarykeys.Count > 0)
                    Attendance.Instance.SyncDataBackground(primarykeys.ToArray());
            };

            List<PackageWorkEventArgs<AttendanceRecordEditor>> packages = worker.Run(editors);
            foreach (PackageWorkEventArgs<AttendanceRecordEditor> each in packages)
            {
                if (each.HasException)
                    throw each.Exception;
            }
        }
    }
}
