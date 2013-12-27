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
    public static class EditMerit
    {
        private static  string INSERT_DEMERIT = "SmartSchool.Student.Discipline.Insert";
        private static  string UPDATE_DEMERIT = "SmartSchool.Student.Discipline.Update";
        private static  string DELETE_DEMERIT = "SmartSchool.Student.Discipline.Delete";

        /// <summary>
        /// 一次儲存一個Editor。
        /// </summary>
        /// <param name="editor"></param>
        internal static void SaveMeritRecordEditor(MeritRecordEditor editor)
        {
            List<MeritRecordEditor> editors = new List<MeritRecordEditor>();
            editors.Add(editor);
            EditMerit.SaveMeritRecordEditors(editors);
        }

        /// <summary>
        /// 批次儲存多個Editor
        /// </summary>
        /// <param name="editors"></param>
        internal static void SaveMeritRecordEditors(IEnumerable<MeritRecordEditor> editors)
        {
            

            string serviceName = "";
            DSXmlHelper helper = null;            

            MultiThreadWorker<MeritRecordEditor> worker = new MultiThreadWorker<MeritRecordEditor>();
            worker.MaxThreads = 3;
            worker.PackageSize = 100;
            worker.PackageWorker += delegate(object sender, PackageWorkEventArgs<MeritRecordEditor> e)
            {
                //1. 對於每一個 editor，               
                foreach (MeritRecordEditor editor in e.List)
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
                foreach (MeritRecordEditor editor in e.List)
                    primarykeys.Add(editor.RefStudentID);

                if (primarykeys.Count > 0)
                    Merit.Instance.SyncDataBackground(primarykeys.ToArray());
            };

            List<PackageWorkEventArgs<MeritRecordEditor>> packages = worker.Run(editors);
            foreach (PackageWorkEventArgs<MeritRecordEditor> each in packages)
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
        private static DSXmlHelper makeInsertHelper(MeritRecordEditor editor)
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
        private static DSXmlHelper makeUpdateHelper(MeritRecordEditor editor)
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
        private static DSXmlHelper makeDeleteHelper(MeritRecordEditor editor)
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
        private static string GetDetailContent(MeritRecordEditor editor)
        {
            DSXmlHelper helper = new DSXmlHelper("Discipline");
            XmlElement element = helper.AddElement("Merit");

            element.SetAttribute("A", editor.MeritA);
            element.SetAttribute("B", editor.MeritB);
            element.SetAttribute("C", editor.MeritC);

            return helper.GetRawXml();
        }

    }
}