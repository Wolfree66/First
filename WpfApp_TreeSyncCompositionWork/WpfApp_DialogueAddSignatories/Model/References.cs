using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Users;

namespace WpfApp_DialogueAddSignatories.Model
{
    internal class References
    {
        static readonly Guid ref_ProjectManagement_Guid = new Guid("86ef25d4-f431-4607-be15-f52b551022ff"); // Справочник "Управление проектами"

        //поля

        private static Reference _ProjectManagementReference;
        private static ReferenceInfo _ProjectManagementReferenceInfo;


        public static ServerConnection Connection
        {
            get { return ServerGateway.Connection; }
        }

        private static Reference GetReference(ref Reference reference, ReferenceInfo referenceInfo)
        {
            if (reference == null)
                reference = referenceInfo.CreateReference();


            return reference;
        }


        static UserReference _UsersReference;
        /// <summary>
        /// Справочник "Группы и пользователи"
        /// </summary>
        public static UserReference UsersReference
        {
            get
            {
                if (_UsersReference == null)
                    _UsersReference = new UserReference(Connection);

                return _UsersReference;
            }
        }

        private static ClassObject _Class_ProjectManagementWork;
        /// <summary>
        /// тип Работа - справочника Управление проектами
        /// </summary>
        public static ClassObject Class_ProjectManagementWork
        {
            get
            {
                if (_Class_ProjectManagementWork == null)
                    _Class_ProjectManagementWork = ProjectManagementReference.Classes.Find(ProjectManagementWork.PM_class_Work_Guid);

                return _Class_ProjectManagementWork;
            }
        }

        /// <summary>
        /// Справочник "Управление проектами"
        /// </summary>
        public static Reference ProjectManagementReference
        {
            get
            {
                if (_ProjectManagementReference == null)
                    _ProjectManagementReference = GetReference(ref _ProjectManagementReference, ProjectManagementReferenceInfo);
                _ProjectManagementReference.Refresh();
                return _ProjectManagementReference;
            }
        }



        /// <summary>
        /// Справочник "Управление проектами"
        /// </summary>
        public static ReferenceInfo ProjectManagementReferenceInfo
        {
            get { return GetReferenceInfo(ref _ProjectManagementReferenceInfo, ref_ProjectManagement_Guid); }
        }



        #region "Используемые ресурсы"

        // Справочник "Используемые ресурсы"
        static readonly Guid ref_UsedResources_Guid = new Guid("3459a8fb-6bca-47ca-971a-1572b684e92e");        //Guid справочника - "Используемые ресурсы"
        static readonly Guid UR_class_NonConsumableResources_Guid = new Guid("8473c817-68fd-479c-a4c3-d6b3b405ea5d"); //Guid типа "Нерасходуемые ресурсы"

        private static Reference _UsedResourcesReference;
        /// <summary>
        /// Справочник "Используемые ресурсы"
        /// </summary>
        public static Reference UsedResources
        {
            get
            {
                if (_UsedResourcesReference == null)
                    _UsedResourcesReference = GetReference(ref _UsedResourcesReference, UsedResourcesReferenceInfo);

                return _UsedResourcesReference;
            }
        }

        private static ReferenceInfo _UsedResourcesReferenceInfo;
        /// <summary>
        /// Справочник "Используемые ресурсы"
        /// </summary>
        private static ReferenceInfo UsedResourcesReferenceInfo
        {
            get { return GetReferenceInfo(ref _UsedResourcesReferenceInfo, ref_UsedResources_Guid); }
        }

        private static TFlex.DOCs.Model.Classes.ClassObject _Class_NonConsumableResources;
        /// <summary>
        /// класс Нерасходуемые ресурсы
        /// </summary>
        public static ClassObject Class_NonConsumableResources
        {
            get
            {
                if (_Class_NonConsumableResources == null)
                    _Class_NonConsumableResources = UsedResources.Classes.Find(UR_class_NonConsumableResources_Guid);

                return _Class_NonConsumableResources;
            }
        }
        #endregion

        #region "Ресурсы"

        // Справочник "Используемые ресурсы"
        static readonly Guid ref_Resources_Guid = new Guid("fe80ab68-01e1-4a95-96cf-602ec877ff19");        //Guid справочника - "Ресурсы"
        private static Reference _ResourcesReference;
        /// <summary>
        /// Справочник "Ресурсы"
        /// </summary>
        public static Reference Resources
        {
            get
            {
                if (_ResourcesReference == null)
                    _ResourcesReference = GetReference(ref _ResourcesReference, ResourcesReferenceInfo);

                return _UsedResourcesReference;
            }
        }

        private static ReferenceInfo _ResourcesReferenceInfo;
        /// <summary>
        /// Справочник "Ресурсы"
        /// </summary>
        private static ReferenceInfo ResourcesReferenceInfo
        {
            get { return GetReferenceInfo(ref _ResourcesReferenceInfo, ref_Resources_Guid); }
        }

        //private static ClassObject _Class_NonConsumableResources;
        ///// <summary>
        ///// класс Нерасходуемые ресурсы
        ///// </summary>
        //public static ClassObject Class_NonConsumableResources
        //{
        //    get
        //    {
        //        if (_Class_NonConsumableResources == null)
        //            _Class_NonConsumableResources = UsedResources.Classes.Find(UR_class_NonConsumableResources_Guid);

        //        return _Class_NonConsumableResources;
        //    }
        //}
        #endregion

        private static ReferenceInfo GetReferenceInfo(ref ReferenceInfo referenceInfo, Guid referenceGuid)
        {
            if (referenceInfo == null)
                referenceInfo = Connection.ReferenceCatalog.Find(referenceGuid);

            return referenceInfo;
        }
    }
}
