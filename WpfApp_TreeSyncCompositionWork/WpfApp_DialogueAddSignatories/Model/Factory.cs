using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using WpfApp_DialogueAddSignatories.Model;

namespace WpfApp_DialogueAddSignatories.Model
{
    static public class Factory
    {


        public static dynamic Create_ProjectManagementWork(ReferenceObject referenceObject)
        {
            ProjectManagementWork projectManagementWork = TryGetFromCache_ProjectManagementWork(referenceObject);
            if (projectManagementWork == null)
            {
                projectManagementWork = new ProjectManagementWork(referenceObject);
                List_ProjectManagementWorks_Cache.Add(projectManagementWork);
            }
            return projectManagementWork;
        }

        private static ProjectManagementWork TryGetFromCache_ProjectManagementWork(ReferenceObject referenceObject)
        {
            List<ProjectManagementWork> listFoundedObject = new List<ProjectManagementWork>();
            foreach (var item in List_ProjectManagementWorks_Cache)
            {
                if (item.ReferenceObject.SystemFields.Id == referenceObject.SystemFields.Id) listFoundedObject.Add(item);
            }
            //var listFoundedDobject = List_ProjectManagementWorks_Cache.Where(on => on.ReferenceObject.SystemFields.Id == referenceObject.SystemFields.Id).ToList();
            if (listFoundedObject == null) return null;
            ProjectManagementWork currentVersion = null;
            foreach (var item in listFoundedObject)
            {
                if (referenceObject.SystemFields.EditDate != item.LastEditDate)
                {
                    List_ProjectManagementWorks_Cache.Remove(item);
                }
                else currentVersion = item;
            }
            return currentVersion;
        }

        internal static ITreeNode CreateProjectTreeItem(ReferenceObject referenceObject)
        {
            //если это работа то создаём работу
            if (referenceObject.Class.IsInherit(References.ProjectManagementReference.Classes.Find(ProjectManagementWork.PM_class_ProjectElement_Guid))) return Create_ProjectManagementWork(referenceObject);
            if (referenceObject.Class.IsInherit(References.ProjectManagementReference.Classes.Find(ProjectFolder.PM_class_ProjectFolder_Guid))) return new ProjectFolder(referenceObject);
            return null;
        }

        //private static OfficialNoteBase TryGetFromCache(ReferenceObject referenceObject)
        //{
        //    return null;
        //    var listFoundedObjects = List_OfficialNotes_Cache.Where(on => on.ReferenceObject.SystemFields.Id == referenceObject.SystemFields.Id).ToList();
        //    if (listFoundedObjects == null) return null;
        //    OfficialNoteBase currentVersion = null;
        //    foreach (var item in listFoundedObjects)
        //    {
        //        if (referenceObject.SystemFields.EditDate != item.LastEditDate)
        //        {
        //            List_OfficialNotes_Cache.Remove(item);
        //        }
        //        else currentVersion = item;
        //    }
        //    return currentVersion;
        //}

        static Factory()
        {

            if (List_ProjectManagementWorks_Cache == null) List_ProjectManagementWorks_Cache = new List<ProjectManagementWork>();
            //if (List_OfficialNotes_Cache == null) List_OfficialNotes_Cache = new List<OfficialNoteBase>();
        }

        static List<ProjectManagementWork> List_ProjectManagementWorks_Cache;
        //static List<OfficialNoteBase> List_OfficialNotes_Cache;
    }
}
