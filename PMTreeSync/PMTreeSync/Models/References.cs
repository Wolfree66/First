using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Users;

namespace TFLEXDocsClasses
{
    /// <summary>
    /// Справочники
    /// ver 1.5
    /// </summary>
    public class References
    {

        public static ServerConnection Connection
        {
            get { return ServerGateway.Connection; }
        }
        #region Справочники и классы
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

        private static Reference _WorkChangeRequestReference;
        /// <summary>
        /// Справочник "Запросы коррекции работ"
        /// </summary>
        public static Reference WorkChangeRequestReference
        {
            get
            {
                if (_WorkChangeRequestReference == null)
                    return GetReference(ref _WorkChangeRequestReference, WorkChangeRequestReferenceInfo);

                return _WorkChangeRequestReference;
            }
        }

        private static ClassObject _Class_PMWorkChangeRequest;
        /// <summary>
        /// класс Запрос коррекции работ
        /// </summary>
        public static ClassObject Class_WorkChangeRequest
        {
            get
            {
                if (_Class_PMWorkChangeRequest == null)
                {
                    Guid CR_class_PMWorkChangeRequest_Guid = new Guid("f79a353c-00da-4248-bd4f-dab6543f95b0"); // Тип "Запросы коррекции работ"
                    _Class_PMWorkChangeRequest = WorkChangeRequestReference.Classes.Find(CR_class_PMWorkChangeRequest_Guid);
                }

                return _Class_PMWorkChangeRequest;
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
                {
                    Guid PM_class_Work_Guid = new Guid("c0bef497-cf64-44a7-9839-a704dc3facb2");
                    _Class_ProjectManagementWork = ProjectManagementReference.Classes.Find(PM_class_Work_Guid);
                }
                return _Class_ProjectManagementWork;
            }
        }

        private static Reference _ProjectManagementReference;
        /// <summary>
        /// Справочник "Управление проектами"
        /// </summary>
        public static Reference ProjectManagementReference
        {
            get
            {
                if (_ProjectManagementReference == null)
                    _ProjectManagementReference = GetReference(ref _ProjectManagementReference, ProjectManagementReferenceInfo);

                return _ProjectManagementReference;
            }
        }


        private static ReferenceInfo _ProjectManagementReferenceInfo;
        /// <summary>
        /// Справочник "Управление проектами"
        /// </summary>
        private static ReferenceInfo ProjectManagementReferenceInfo
        {
            get
            {
                Guid ref_ProjectManagement_Guid = new Guid("86ef25d4-f431-4607-be15-f52b551022ff"); // Справочник "Управление проектами"
                return GetReferenceInfo(ref _ProjectManagementReferenceInfo, ref_ProjectManagement_Guid);
            }
        }

        private static ReferenceInfo _WorkChangeRequestReferenceInfo;
        /// <summary>
        /// Справочник "Запросы коррекции работ"
        /// </summary>
        internal static ReferenceInfo WorkChangeRequestReferenceInfo
        {
            get
            {
                Guid WorkChangeRequests_ref_Guid = new Guid("9387cd96-d3cf-4cb6-997d-c4f4e59f8a21"); // Справочник "Запросы коррекции работ"
                return GetReferenceInfo(ref _WorkChangeRequestReferenceInfo, WorkChangeRequests_ref_Guid);
            }
        }

        static FileReference _FileReference;
        /// <summary>
        /// Справочник "Файлы"
        /// </summary>
        public static FileReference FileReference
        {
            get
            {
                if (_FileReference == null)
                    _FileReference = new FileReference(Connection);

                return _FileReference;
            }
        }
        #endregion
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

        private static ClassObject _Class_NonConsumableResources;
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

        // Справочник "Ресурсы"
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

                return _ResourcesReference;
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
        #endregion

        #region "Регистрационно-Контрольные карточки (Канцелярия)"

        static readonly Guid ref_RegistryControlCards_Guid = new Guid("80831dc7-9740-437c-b461-8151c9f85e59");
        private static Reference _RegistryControlCardsReference;
        /// <summary>
        /// Справочник "Регистрационно-Контрольные карточки"
        /// </summary>
        public static Reference RegistryControlCards
        {
            get
            {
                if (_RegistryControlCardsReference == null)
                    _RegistryControlCardsReference = GetReference(ref _RegistryControlCardsReference, RegistryControlCardsReferenceInfo);

                _RegistryControlCardsReference.Refresh();
                return _RegistryControlCardsReference;
            }
        }

        private static ReferenceInfo _RegistryControlCardsReferenceInfo;
        /// <summary>
        /// Справочник "Регистрационно-Контрольные карточки"
        /// </summary>
        private static ReferenceInfo RegistryControlCardsReferenceInfo
        {
            get { return GetReferenceInfo(ref _RegistryControlCardsReferenceInfo, ref_RegistryControlCards_Guid); }
        }
        #endregion

        #region "Статистика загрузки ресурсов"

        // Справочник "Статистика загрузки ресурсов"
        static readonly Guid ref_StatisticConsumptionResource_Guid = new Guid("fb63e9e9-71b3-4afc-aed5-8312a732a60a");        //Guid справочника - "Статистика загрузки ресурсов"
        private static Reference _StatisticConsumptionResourceReference;
        /// <summary>
        /// Справочник "Статистика загрузки ресурсов"
        /// </summary>
        public static Reference StatisticConsumptionResourceReference
        {
            get
            {
                if (_StatisticConsumptionResourceReference == null)
                    _StatisticConsumptionResourceReference = GetReference(ref _StatisticConsumptionResourceReference, StatisticConsumptionResourceReferenceInfo);

                return _StatisticConsumptionResourceReference;
            }
        }

        private static ReferenceInfo _StatisticConsumptionResourceReferenceInfo;
        /// <summary>
        /// Справочник "Статистика загрузки ресурсов"
        /// </summary>
        private static ReferenceInfo StatisticConsumptionResourceReferenceInfo
        {
            get { return GetReferenceInfo(ref _StatisticConsumptionResourceReferenceInfo, ref_StatisticConsumptionResource_Guid); }
        }


        static readonly Guid SCR_class_StatisticConsumptionResource_Guid = new Guid("41606f1a-16c5-4df4-b507-360590ebfe95"); //Guid типа "Загрузка ресурса"
        static ClassObject _Class_StatisticConsumptionResource;
        /// <summary>
        /// класс Нерасходуемые ресурсы
        /// </summary>
        public static ClassObject Class_StatisticConsumptionResource
        {
            get
            {
                if (_Class_StatisticConsumptionResource == null)
                    _Class_StatisticConsumptionResource = StatisticConsumptionResourceReference.Classes.Find(SCR_class_StatisticConsumptionResource_Guid);

                return _Class_StatisticConsumptionResource;
            }
        }
        #endregion

        #region Регистрационно-контрольные карточки

        static Reference _RegistryControlCardReference;

        /// <summary>
        /// Справочник "Регистрационно-контрольные карточки"
        /// </summary>
        private Reference RegistryControlCardReference
        {
            get
            {
                if (_RegistryControlCardReference == null)

                    return GetReference(ref _RegistryControlCardReference, RegistryControlCardReferenceInfo);
                _RegistryControlCardReference.Refresh();
                return _RegistryControlCardReference;
            }
        }

        private static Guid reference_RCC_Guid = new Guid("80831dc7-9740-437c-b461-8151c9f85e59");    //Guid справочника - "Регистрационно-контрольные карточки"
        static ReferenceInfo _RCCReferenceInfo;
        private ReferenceInfo RegistryControlCardReferenceInfo
        {
            get { return GetReferenceInfo(ref _RCCReferenceInfo, reference_RCC_Guid); }
        }
        #endregion



        private static ReferenceInfo GetReferenceInfo(ref ReferenceInfo referenceInfo, Guid referenceGuid)
        {
            if (referenceInfo == null)
                referenceInfo = Connection.ReferenceCatalog.Find(referenceGuid);

            return referenceInfo;
        }
    }
}
