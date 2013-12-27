using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JHSchool.Behavior.Editor;
using Framework;
using FISCA.DSAUtil;
using JHSchool.Editor;
using System.Xml;
using FISCA.Authentication;

namespace JHSchool.Behavior.Feature
{
    public static class EditDemerit
    {
        private static  string INSERT_DEMERIT = "SmartSchool.Student.Discipline.Insert";
        private static  string UPDATE_DEMERIT = "SmartSchool.Student.Discipline.Update";
        private static  string DELETE_DEMERIT = "SmartSchool.Student.Discipline.Delete";


        /// <summary>
        /// 儲存一個銷過的 Editor
        /// </summary>
        /// <param name="editor"></param>
        internal static void SaveClearDemeritRecordEditor(ClearDemeritRecordEditor editor)
        {
            List<ClearDemeritRecordEditor> editors = new List<ClearDemeritRecordEditor>();
            editors.Add(editor);
            SaveClearDemeritRecordEditors(editors);
        }

        /// <summary>
        /// 批次儲存銷過的 Editor
        /// </summary>
        /// <param name="editors"></param>
        internal static void SaveClearDemeritRecordEditors(IEnumerable<ClearDemeritRecordEditor> editors)
        {
            string serviceName = UPDATE_DEMERIT;

            MultiThreadWorker<ClearDemeritRecordEditor> worker = new MultiThreadWorker<ClearDemeritRecordEditor>();
            worker.MaxThreads = 3;
            worker.PackageSize = 100;
            worker.PackageWorker += delegate(object sender, PackageWorkEventArgs<ClearDemeritRecordEditor> e)
            {
                //1. 對於每一個 editor，               
                foreach (ClearDemeritRecordEditor editor in e.List)
                {
                    DSXmlHelper helper = makeUpdateHelper(editor);
                    DSAServices.CallService(serviceName, new DSRequest(helper.BaseElement));
                }

                //4. 呼叫 CacheManager ，重新取得這些學生的懲戒記錄。
                List<string> primarykeys = new List<string>();
                foreach (ClearDemeritRecordEditor editor in e.List)
                {
                    if (!primarykeys.Contains(editor.RefStudentID))     //傳入 Demerit.Instance.SyncDataBackground 的 key不能重複。
                        primarykeys.Add(editor.RefStudentID);
                }

                if (primarykeys.Count > 0)
                    Demerit.Instance.SyncDataBackground(primarykeys.ToArray());

            };

            List<PackageWorkEventArgs<ClearDemeritRecordEditor>> packages = worker.Run(editors);
            foreach (PackageWorkEventArgs<ClearDemeritRecordEditor> each in packages)
            {
                if (each.HasException)
                    throw each.Exception;
            }

        }



        /// <summary>
        /// 一次儲存一個Editor。
        /// </summary>
        /// <param name="editor"></param>
        internal static void SaveDemeritRecordEditor(DemeritRecordEditor editor)
        {
            List<DemeritRecordEditor> editors = new List<DemeritRecordEditor>();
            editors.Add(editor);
            EditDemerit.SaveDemeritRecordEditors(editors);
        }

        /// <summary>
        /// 批次儲存多個Editor
        /// </summary>
        /// <param name="editors"></param>
        internal static void SaveDemeritRecordEditors(IEnumerable<DemeritRecordEditor> editors)
        {
            string serviceName = "";
            DSXmlHelper helper = null;            

            MultiThreadWorker<DemeritRecordEditor> worker = new MultiThreadWorker<DemeritRecordEditor>();
            worker.MaxThreads = 3;
            worker.PackageSize = 100;
            worker.PackageWorker += delegate(object sender, PackageWorkEventArgs<DemeritRecordEditor> e)
            {
                //1. 對於每一個 editor，               
                foreach (DemeritRecordEditor editor in e.List)
                {
                    //2. 判斷狀態是新增，修改或是刪除
                    //2.1 決定相對的 服務名稱
                    //2.2 建立相對該狀態的 Request Document
                    switch (editor.EditorStatus)
                    {
                        case EditorStatus.Insert:
                            serviceName = INSERT_DEMERIT;
                            helper = makeInsertHelper(editor);
                            break;
                        case EditorStatus.Update:
                            serviceName = UPDATE_DEMERIT;
                            helper = makeUpdateHelper(editor);
                            break;
                        case EditorStatus.Delete:
                            serviceName = DELETE_DEMERIT;
                            helper = makeDeleteHelper(editor);
                            break;
                        default :
                            continue;
                    }

                    //3. 呼叫相關的Services
                    DSAServices.CallService(serviceName , new DSRequest(helper.BaseElement));
                }               
                
                //4. 呼叫 CacheManager ，重新取得這些學生的懲戒記錄。
                List<string> primarykeys = new List<string>();
                foreach (DemeritRecordEditor editor in e.List)
                    primarykeys.Add(editor.RefStudentID);

                if (primarykeys.Count > 0)
                    Demerit.Instance.SyncDataBackground(primarykeys.ToArray());
            };

            List<PackageWorkEventArgs<DemeritRecordEditor>> packages = worker.Run(editors);
            foreach (PackageWorkEventArgs<DemeritRecordEditor> each in packages)
            {
                if (each.HasException)
                    throw each.Exception;
            }
        }

        /// <summary>
        /// 組合出 新增所需要的Request Document格式
        /// </summary>
        /// <param name="editor"></param>
        /// <returns></returns>
        private static DSXmlHelper makeInsertHelper(DemeritRecordEditor editor)
        {
            DSXmlHelper helper = new DSXmlHelper("InsertRequest");
            helper.AddElement("Discipline");
            helper.AddElement("Discipline", "RefStudentID", editor.RefStudentID);
            helper.AddElement("Discipline", "SchoolYear", editor.SchoolYear);
            helper.AddElement("Discipline", "Semester", editor.Semester);
            helper.AddElement("Discipline", "GradeYear", (editor.GradeYear == "未分年級") ? "" : editor.GradeYear);
            helper.AddElement("Discipline", "OccurDate", editor.OccurDate);
            helper.AddElement("Discipline", "Reason", editor.Reason);
            helper.AddElement("Discipline", "RegisterDate", editor.RegisterDate);
            helper.AddElement("Discipline", "MeritFlag", editor.MeritFlag);
            helper.AddElement("Discipline", "Type", "1");
            helper.AddElement("Discipline", "Detail", GetDetailContent(editor), true);

            return helper;
        }

        /// <summary>
        /// 組合出 修改所需要的Request Document格式
        /// </summary>
        /// <param name="editor"></param>
        /// <returns></returns>
        private static DSXmlHelper makeUpdateHelper(DemeritRecordEditor editor)
        {
            DSXmlHelper helper = new DSXmlHelper("UpdateRequest");

            helper.AddElement("Discipline");
            helper.AddElement("Discipline", "Field");
            helper.AddElement("Discipline/Field", "RefStudentID", editor.RefStudentID);
            helper.AddElement("Discipline/Field", "SchoolYear", editor.SchoolYear);
            helper.AddElement("Discipline/Field", "Semester", editor.Semester);
            helper.AddElement("Discipline/Field", "GradeYear", (editor.GradeYear == "未分年級") ? "" : editor.GradeYear);
            helper.AddElement("Discipline/Field", "OccurDate", editor.OccurDate);
            helper.AddElement("Discipline/Field", "Reason", editor.Reason);
            helper.AddElement("Discipline/Field", "RegisterDate", editor.RegisterDate);
            helper.AddElement("Discipline/Field", "MeritFlag", editor.MeritFlag);
            helper.AddElement("Discipline/Field", "Type", "1");
            helper.AddElement("Discipline/Field", "Detail", GetDetailContent(editor), true);
            helper.AddElement("Discipline", "Condition");
            helper.AddElement("Discipline/Condition", "ID", editor.ID);

            return helper;
        }

        /// <summary>
        /// 組合出 刪除所需要的Request Document格式
        /// </summary>
        /// <param name="editor"></param>
        /// <returns></returns>
        private static DSXmlHelper makeDeleteHelper(DemeritRecordEditor editor)
        {
            DSXmlHelper helper = new DSXmlHelper("DeleteRequest");

            helper.AddElement("Discipline");
            helper.AddElement("Discipline", "ID",editor.ID);

            return helper;
        }

        /// <summary>
        /// 組合出 Detail 節點的內容。新增和修改的 Request Doc 都會用到。
        /// </summary>
        /// <param name="editor"></param>
        /// <returns></returns>
        private static string GetDetailContent(DemeritRecordEditor editor)
        {
            DSXmlHelper helper = new DSXmlHelper("Discipline");
            XmlElement element = helper.AddElement("Demerit");

            element.SetAttribute("A", editor.DemeritA);
            element.SetAttribute("B", editor.DemeritB);
            element.SetAttribute("C", editor.DemeritC);            
            
            bool isClearDemerit = (editor.GetType() == Type.GetType("JHSchool.Behavior.Editor.ClearDemeritRecordEditor"));

            if (editor.IsCleared)
            {
                element.SetAttribute("Cleared", editor.Cleared);
                element.SetAttribute("ClearDate", editor.ClearDate);
                element.SetAttribute("ClearReason", editor.ClearReason);
            }
            else if (isClearDemerit)
            {
                ClearDemeritRecordEditor cdr = (ClearDemeritRecordEditor)editor;
                element.SetAttribute("Cleared", cdr.Cleared );
                element.SetAttribute("ClearDate", cdr.ClearDate);
                element.SetAttribute("ClearReason", cdr.ClearReason );
            }
            else
            {
                element.SetAttribute("Cleared", "");
                element.SetAttribute("ClearDate", "");
                element.SetAttribute("ClearReason", "");
            }

            return helper.GetRawXml();
        }

    }
}
