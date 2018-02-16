
using System;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Common;
using System.Collections.Generic;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.Access;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using System.Text;
using System.Linq;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Utils;
using TFlex.DOCs.Model.Signatures;
using TFlex.DOCs.Model.References.Calendar;
using TFlex.DOCs.Model.References.Files;
using DinamikaGuids;

namespace Dinamika
{

    /// <summary>
    /// Методы Динамики
    /// </summary>
    public class ExtensionDinamika
    {
        public ExtensionDinamika(ServerConnection connection = null, string nameServer = "TFLEX")
        {
            if (connection != null)
                _connection = connection;
            this.nameServer = nameServer;

        }

        public ExtensionDinamika() : this(null, "TFLEX")
        {

        }

        #region Поля

        private string nameServer;

        private ServerConnection _connection;

        private User _currentUser;

        private Reference _projectResourcesReference;

        private Reference _projectManagementReference;

        private Reference _officeNotes;

        private Reference _fileReference;

        private UserReference _userReference;

        private Reference _dependenciesReference;

        private Reference _resourceUsed;

        private Reference _errandControl;

        private ReferenceInfo _dependenciesReferenceInfo;

        private ReferenceInfo _fileReferenceReferenceInfo;

        private ReferenceInfo _resourceUsedInfo;

        private ReferenceInfo _projectResourcesReferenceInfo;

        private ReferenceInfo _userReferenceInfo;

        private ReferenceInfo _projectManagementReferenceInfo;

        private ReferenceInfo _errandControlReferenceInfo;

        private ReferenceInfo _officeNotesReferenceInfo;

        /// <summary>
        /// Структрура гуидов
        /// </summary>
        public Guids Guids = new Guids();

        //private  FileObject DinamikaDll;

        //private  FileObject DinamikaDialogExtensionDll;

        #endregion


        #region Свойства

        /// <summary>
        /// Текущий пользователь
        /// </summary>
        public User currentUser
        {
            get
            {
                if (_currentUser == null)
                {
                    _currentUser = (User)Connection.ClientView.GetUser();
                    return _currentUser;
                }
                else
                    return _currentUser;


            }
        }

        /// <summary>
        /// servConnection
        /// </summary>
        public ServerConnection Connection
        {
            get
            {
                if (_connection == null)
                {
                    try
                    {
                        _connection = ServerGateway.Connection;
                    }
                    catch (InvalidOperationException)
                    {
                        try
                        {
                            Connect();
                            _connection = ServerGateway.Connection;
                        }
                        catch { }
                    }

                    return _connection;
                }
                else
                    return _connection;

            }
        }

        /// <summary>
        /// servConnection
        /// </summary>
        public void Connect()
        {
            if (ServerGateway.IsConnected)
            {
                _connection = ServerGateway.Connection;
                return;
            }

            ServerGateway.Connect(nameServer);
        }


        #region Описания справочников

        /// <summary>
        /// Описание спр. - "Файлы"
        /// </summary>
        public ReferenceInfo ErrandControlReferenceInfo
        {
            get
            {
                if (_errandControlReferenceInfo == null)
                    _errandControlReferenceInfo = GetReferenceInfo(Guids.ErrandControl.Ref_Guid);

                return _errandControlReferenceInfo;
            }
        }

        /// <summary>
        /// Описание спр. - "Файлы"
        /// </summary>
        public ReferenceInfo FileReferenceReferenceInfo
        {
            get
            {
                if (_fileReferenceReferenceInfo == null)
                    _fileReferenceReferenceInfo = GetReferenceInfo(Guids.FileReference.Ref_Guid);

                return _fileReferenceReferenceInfo;
            }
        }
        /// <summary>
        /// Описание спр. - "Служебные записки"
        /// </summary>
        public ReferenceInfo OfficeNotesReferenceInfo
        {
            get
            {
                if (_officeNotesReferenceInfo == null)
                    _officeNotesReferenceInfo = GetReferenceInfo(Guids.OfficeNotes.Ref_Guid);

                return _officeNotesReferenceInfo;
            }
        }
        /// <summary>
        /// Описание спр. - "Ресурсы"
        /// </summary>
        public ReferenceInfo ProjectResourcesReferenceInfo
        {
            get
            {

                if (_projectResourcesReferenceInfo == null)
                    _projectResourcesReferenceInfo = GetReferenceInfo(Guids.ProjectResourcesReference.Ref_Guid);

                return _projectResourcesReferenceInfo;
            }
        }

        /// <summary>
        /// Описание спр. - "Используемые ресурсы"
        /// </summary>
        public ReferenceInfo ResourceUsedReferenceInfo
        {
            get
            {

                if (_resourceUsedInfo == null)
                    _resourceUsedInfo = GetReferenceInfo(Guids.ResourceUsedReference.Ref_Guid);

                return _resourceUsedInfo;
            }
        }

        /// <summary>
        /// Описание спр. - "Зависимости проектов"
        /// </summary>
        public ReferenceInfo DependenciesReferenceInfo
        {
            get
            {

                if (_dependenciesReferenceInfo == null)
                    _dependenciesReferenceInfo = GetReferenceInfo(Guids.DependenciesReference.Ref_Guid);

                return _dependenciesReferenceInfo;
            }
        }


        /// <summary>
        ///  Описания справочника - "Группы и пользователи"
        /// </summary>
        public ReferenceInfo UserReferenceInfo
        {
            get
            {
                if (_userReferenceInfo == null)
                    _userReferenceInfo = GetReferenceInfo(Guids.UserReference.Ref_Guid);

                return _userReferenceInfo;
            }
        }

        /// <summary>
        ///  Описания справочника - "Управление проектами"
        /// </summary>
        public ReferenceInfo ProjectManagementReferenceInfo
        {
            get
            {

                if (_projectManagementReferenceInfo == null)
                    _projectManagementReferenceInfo = GetReferenceInfo(Guids.ProjectManagementReference.Ref_Guid);

                return _projectManagementReferenceInfo;
            }
        }
        #endregion


        #region Справочники

        /// <summary>
        /// Контроль поручений
        /// </summary>
        public Reference ErrandControl
        {
            get
            {
                if (_errandControl == null)
                    _errandControl = GetReference(ErrandControlReferenceInfo);

                return _errandControl;
            }
        }

        /// <summary>
        /// Файлы
        /// </summary>
        public Reference FileReference
        {
            get
            {
                if (_fileReference == null)
                    _fileReference = GetReference(FileReferenceReferenceInfo);

                return _fileReference;
            }
        }

        /// <summary>
        /// Служебные записки
        /// </summary>
        public Reference OfficeNotes
        {
            get
            {
                if (_officeNotes == null)
                    _officeNotes = GetReference(OfficeNotesReferenceInfo);

                return _officeNotes;
            }
        }
        /// <summary>
        /// Ресурсы
        /// </summary>
        public Reference ProjectResourcesReference
        {
            get
            {
                if (_projectResourcesReference == null)
                    _projectResourcesReference = GetReference(ProjectResourcesReferenceInfo);

                return _projectResourcesReference;
            }
        }


        /// <summary>
        /// Управление проектами
        /// </summary>
        public Reference ProjectManagementReference
        {
            get
            {
                if (_projectManagementReference == null)
                    _projectManagementReference = GetReference(ProjectManagementReferenceInfo);

                return _projectManagementReference;
            }
        }

        /// <summary>
        /// Группы и пользователи
        /// </summary>
        public UserReference UserReference
        {
            get
            {
                if (_userReference == null)
                    _userReference = GetReference(UserReferenceInfo) as UserReference;

                return _userReference;
            }
        }

        /// <summary>
        /// Зависимости проектов
        /// </summary>
        public Reference DependenciesReference
        {
            get
            {
                if (_dependenciesReference == null)
                    _dependenciesReference = GetReference(DependenciesReferenceInfo);

                return _dependenciesReference;
            }
        }

        /// <summary>
        /// Используемые ресурсы
        /// </summary>
        public Reference ResourceUsed
        {
            get
            {
                if (_resourceUsed == null)
                    _resourceUsed = GetReference(ResourceUsedReferenceInfo);

                return _resourceUsed;
            }
        }

        #endregion

        #endregion

        #region расширяющие методы УП

        /// <summary>
        /// Список проектов папки "предварительное согласование", которые нужно подписать
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public List<ReferenceObject> ПланыНаПодпись_ПапкаПредСогл(ReferenceObject user)
        {

            List<ReferenceObject> ПланыСогл_НаПодпись = new List<ReferenceObject>();

            List<ReferenceObject> ПланыСогл = new List<ReferenceObject>();

            ПланыСогл.AddRange(ProjectManagementReference.Find(Guids.ProjectManagementReference.TypeProjectElement.Objects.FolderPreliminaryCoordination_Guid)
                .Children.RecursiveLoad().Where(obj => obj.Class.IsInherit(Guids.ProjectManagementReference.TypeProjectElement.TypeProject_Guid)));

            // int i = ПланыПредварСогл.Count;

            foreach (var План in ПланыСогл)
            {
                foreach (Signature sign in План.Signatures)
                {
                    //MessageBox.Show(sign.SignatureObjectType.Name);
                    if (sign.UserName == user[Guids.UserReference.Name_Guid].GetString() && sign.SignatureDate == null /*&& sign.Actual == true*/)
                    {
                        ПланыСогл_НаПодпись.Add(План);
                    }
                }
            }

            return ПланыСогл_НаПодпись = ПланыСогл_НаПодпись.Distinct().ToList();
        }

        /// <summary>
        /// Список проектов папки "согласование планов", которые нужно подписать
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public List<ReferenceObject> ПланыНаПодпись_ПапкаСогл(ReferenceObject user)
        {
            List<ReferenceObject> ПланыПредСогл_НаПодпись = new List<ReferenceObject>();

            List<ReferenceObject> ПланыПредварСогл = new List<ReferenceObject>();

            ПланыПредварСогл.AddRange(ProjectManagementReference.Find(Guids.ProjectManagementReference.TypeProjectElement.Objects.FolderCoordinationPlans_Guid)
                .Children.RecursiveLoad().Where(obj => obj.Class.IsInherit(Guids.ProjectManagementReference.TypeProjectElement.TypeProject_Guid)));

            // int i = ПланыПредварСогл.Count;

            foreach (var План in ПланыПредварСогл)
            {
                foreach (Signature sign in План.Signatures)
                {
                    //MessageBox.Show(sign.SignatureObjectType.Name);
                    if (sign.UserName == user[Guids.UserReference.Name_Guid].GetString() && sign.SignatureDate == null /*&& sign.Actual == true*/)
                    {
                        ПланыПредСогл_НаПодпись.Add(План);
                    }
                }
            }

            return ПланыПредСогл_НаПодпись = ПланыПредСогл_НаПодпись.Distinct().ToList();
        }

        /// <summary>
        /// Добавлятет доступ на элемент проекта, если его нет.
        /// 0 - только чтение
        /// 1 - Без ограничений
        /// 2 - параметры
        /// 4 - ресурсы
        /// 8 - детализация, другие значения цифр см. в АСУ
        /// </summary>
        /// <param name="projectElement">Элемент проекта, на который нужно назначить метод</param>
        /// <param name="user">Кому дать доступ</param>
        /// <param name="Value">Какой дать доступ</param>
        /// <returns></returns>
        public bool ДобавитьДоступНаЭлементПроекта(ReferenceObject projectElement, UserReferenceObject user, int Value)
        {

            if (projectElement == null || user == null)
                return false;

            int i = 0;

            foreach (ReferenceObject Access in projectElement.GetObjects(Guids.ProjectManagementReference.TypeProjectElement.TypeAccesses_AccessesList_Guid))//перебор объектов
            {
                if (Access[Guids.ProjectManagementReference.TypeProjectElement.TypeAccesses_AccessesList.Accesses_TypeParameters_User_Guid].GetString() == user.FullName.ToString()
                  && (int)Access[Guids.ProjectManagementReference.TypeProjectElement.TypeAccesses_AccessesList.Accesses_TypeParameters_Access_Guid].Value == Value
                  && Access[Guids.ProjectManagementReference.TypeProjectElement.TypeAccesses_AccessesList.Accesses_TypeParameters_UserGuid_Guid].GetString() == user.SystemFields.Guid.ToString())//Параметры
                {
                    i++;
                }
            }

            if (i == 0)
            {
                if (МожноЛиРедактироватьОбъект(projectElement) == false) return false;

                projectElement.BeginChanges();

                ReferenceObject newAccess = projectElement
                    .CreateListObject(Guids.ProjectManagementReference.TypeProjectElement.TypeAccesses_AccessesList_Guid,
                    Guids.ProjectManagementReference.TypeProjectElement.TypeAccesses_AccessesList.Accesses_TypeAccess_Guid);

                newAccess[Guids.ProjectManagementReference.TypeProjectElement.TypeAccesses_AccessesList.Accesses_TypeParameters_User_Guid].Value = user.FullName.ToString();//Параметр пользователь
                newAccess[Guids.ProjectManagementReference.TypeProjectElement.TypeAccesses_AccessesList.Accesses_TypeParameters_Access_Guid].Value = Value;//Параметры
                newAccess[Guids.ProjectManagementReference.TypeProjectElement.TypeAccesses_AccessesList.Accesses_TypeParameters_UserGuid_Guid].Value = user.SystemFields.Guid;//Параметры

                newAccess.EndChanges();

                projectElement.EndChanges();

                return true;
            }
            return true;
        }

        /// <summary>
        /// Удаляет указанный доступ с элемента проекта
        /// 0 - только чтение
        /// 1 - Без ограничений
        /// 2 - параметры
        /// 4 - ресурсы
        /// 8 - детализация, другие значения цифр см. в АСУ
        /// </summary>
        /// <param name="projectElement">Элемент проекта, с которого нужно удалить доступ</param>
        /// <param name="user">Кому удалить доступ</param>
        /// <param name="Value">Какой удалить доступ</param>
        /// <returns></returns>
        public bool УдалитьДоступСЭлементаПроекта(ReferenceObject projectElement, User user, int Value)
        {
            if (projectElement == null || user == null)
                return false;

            if (МожноЛиРедактироватьОбъект(projectElement) == false) return false;

            projectElement.BeginChanges();

            foreach (ReferenceObject Access in projectElement.GetObjects(Guids.ProjectManagementReference.TypeProjectElement.TypeAccesses_AccessesList_Guid))//перебор объектов
            {

                if (Access[Guids.ProjectManagementReference.TypeProjectElement.TypeAccesses_AccessesList.Accesses_TypeParameters_User_Guid].GetString() == user.FullName.ToString()
                  && (int)Access[Guids.ProjectManagementReference.TypeProjectElement.TypeAccesses_AccessesList.Accesses_TypeParameters_Access_Guid].Value == Value
                  && Access[Guids.ProjectManagementReference.TypeProjectElement.TypeAccesses_AccessesList.Accesses_TypeParameters_UserGuid_Guid].GetString() == user.SystemFields.Guid.ToString())//Параметры
                {
                    Access.Delete();
                }
            }

            projectElement.EndChanges();

            //projectElement.Reload();

            return true;
        }

        /// <summary>
        /// Удаляет все доступы с элемента проекта
        /// </summary>
        /// <param name="projectElement">Элемент проекта, с которого нужно удалить доступы</param>
        /// <returns></returns>
        public bool УдалитьВсеДоступыСЭлементаПроекта(ReferenceObject projectElement)
        {
            if (projectElement == null)
                return false;

            if (!МожноЛиРедактироватьОбъект(projectElement)) return false;

            projectElement.BeginChanges();

            foreach (ReferenceObject Access in projectElement.GetObjects(Guids.ProjectManagementReference.TypeProjectElement.TypeAccesses_AccessesList_Guid))//перебор объектов
            {
                Access.Delete();
            }

            projectElement.EndChanges();

            //projectElement.Reload();

            return true;
        }




        /// <summary>
        /// возвращает список зависимостей типа синхронизация для работы, без вызова менеджера проекта
        ///(является владельцем true = укрупнение)
        /// </summary>
        public IEnumerable<ReferenceObject> СписокЗависимостейЭлементаПроектаТипаСинхронизация(ReferenceObject projectElement, bool? являетсяВладельцем = null)
        {
            if (projectElement == null)
                return Enumerable.Empty<ReferenceObject>();

            Guid guidForSearch = projectElement.SystemFields.Guid;

            Filter filter = new Filter(DependenciesReferenceInfo);

            Reference dependenciesReference = DependenciesReferenceInfo.CreateReference();

            // Условие: тип зависимости "синхронизация"
            ReferenceObjectTerm termSyncType = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
            // устанавливаем параметр
            termSyncType.Path.AddParameter(dependenciesReference.ParameterGroup.OneToOneParameters.Find(Guids.DependenciesReference.DependencyType_Guid));
            // устанавливаем оператор сравнения
            termSyncType.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termSyncType.Value = 5;

            // Условие: в названии содержится "является укрупнением"
            //ReferenceObjectTerm termMasterWork = new ReferenceObjectTerm(LogicalOperator.Or);

            ReferenceObjectTerm termMasterWork = new ReferenceObjectTerm(filter.Terms);
            // устанавливаем параметр
            termMasterWork.Path.AddParameter(dependenciesReference.ParameterGroup.OneToOneParameters.Find(Guids.DependenciesReference.Work1_Guid));
            // устанавливаем оператор сравнения
            termMasterWork.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termMasterWork.Value = guidForSearch.ToString();
            // Условие: в названии содержится "является детализацией"
            ReferenceObjectTerm termSlaveWork = new ReferenceObjectTerm(filter.Terms, LogicalOperator.Or);
            //ReferenceObjectTerm term1 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.Or);
            // устанавливаем параметр
            termSlaveWork.Path.AddParameter(dependenciesReference.ParameterGroup.OneToOneParameters.Find(Guids.DependenciesReference.Work2_Guid));
            // устанавливаем оператор сравнения
            termSlaveWork.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termSlaveWork.Value = guidForSearch.ToString();

            // Группируем условия в отдельную группу (другими словами добавляем скобки)
            TermGroup group1 = filter.Terms.GroupTerms(new Term[] { termMasterWork, termSlaveWork });

            //редактируем при необходимости
            if (являетсяВладельцем != null && (bool)являетсяВладельцем) group1.Remove(termSlaveWork);
            if (!являетсяВладельцем != null && !((bool)являетсяВладельцем)) group1.Remove(termMasterWork);

            //MessageBox.Show(filter.ToString());
            //Получаем список объектов, в качестве условия поиска – сформированный фильтр
            List<ReferenceObject> listObj = dependenciesReference.Find(filter);

            return listObj;
        }

        /// <summary>
        /// Метод копирования используемых ресурсов
        /// </summary>
        /// <param name="owner">Элемент проекта в который копируем используемые ресурсы</param>
        /// <param name="whence">Элемент проекта с которого копируем используемые ресурс</param>
        /// <param name="PlanningSpaceForNewRes_Guid">Пространство планирования для новых ресурсов</param>
        ///   /// <param name="PlanningSpaceForCheck_Guid">Копировать ресурсы только из этого пространства планирования</param>
        /// <returns></returns>
        public bool СкопироватьИспользуемыеРесурсы_изЭлементаПроекта_вЭлементПроекта(ReferenceObject owner, ReferenceObject whence
            , Guid? PlanningSpaceForNewRes_Guid = null, Guid? PlanningSpaceForCheck_Guid = null)
        {

            if (owner == null || whence == null) return false;

            //получение списка используемых ресурсов с детализации
            var usedResources = GetLinkedUsedResources(whence)
                .Where(UseRes =>
                {
                    ReferenceObject NonExpendableResource = null;

            //Получаем Ресурс из справочника Ресурсы
            NonExpendableResource = GetResourcesLink(UseRes);

                    if (NonExpendableResource == null)
                    {
                //MessageBox.Show("NonExpendableResource == null");
                return false;
                    }

                    if (PlanningSpaceForCheck_Guid != null && PlanningSpaceForCheck_Guid != GetPlanningSpaceUsedResource(UseRes))
                    {
                        return false;
                    }
            //Проверка ресурса на тип
            switch (NonExpendableResource.Class.Name)
                    {
                        case "Оборудование и оснастка":
                            { return false; }
                        case "Оснащение":
                            { return false; }
                        case "Комплектующие":
                            { return false; }
                        case "Ресурсы материалов":
                            { return false; }
                    }

                    return true;
                }).ToList();

            if (usedResources.Count() == 0)
                return true;

            var result = new List<ReferenceObject>(usedResources.Count);

            //цикл копирования используемых ресурсов с детализации, в справочник используемыее ресурсы
            foreach (var usedResource in usedResources)
            {

                //Здесь дописать копирование нужных параметров и связей
                var newResourceUsed = usedResource.CreateCopy(usedResource.Class);

                // var newResourceUsed = ResourceUsed.CreateReferenceObject(null, ResourceUsed.Classes.Find(Dinamika.Guids.ResourceUsedReference.TypeNonExpendableResourcesUsedReferenceObject_Guid));
                /*
                //Получаем Ресурс из справочника Ресурсы
                var NonExpendableResource = usedResource.GetObject(Guids.ResourceUsedReference.Links.ResourcesLink_Guid);

                newResourceUsed[Guids.ResourceUsedReference.Name_Guid].Value = usedResource[Guids.ResourceUsedReference.Name_Guid].Value;

                if (NonExpendableResource != null)
                    newResourceUsed.SetLinkedObject(Guids.ResourceUsedReference.Links.ResourcesLink_Guid, NonExpendableResource);

                newResourceUsed[Guids.ResourceUsedReference.Number_Guid].Value = usedResource[Guids.ResourceUsedReference.Number_Guid].Value;


                newResourceUsed[Guids.ResourceUsedReference.BeginDate_Guid].Value = usedResource[Guids.ResourceUsedReference.BeginDate_Guid].Value;
                newResourceUsed[Guids.ResourceUsedReference.EndDate_Guid].Value = usedResource[Guids.ResourceUsedReference.EndDate_Guid].Value;

                newResourceUsed[Guids.ResourceUsedReference.FactValue_Guid].Value = usedResource[Guids.ResourceUsedReference.FactValue_Guid].Value;
                newResourceUsed[Guids.ResourceUsedReference.FixNumber_Guid].Value = usedResource[Guids.ResourceUsedReference.FixNumber_Guid].Value;
                */

                if (PlanningSpaceForNewRes_Guid != null)
                    if (!newResourceUsed.Changing)
                        newResourceUsed.BeginChanges();
                {
                    newResourceUsed[Guids.ResourceUsedReference.PlanningSpace_Guid].Value = PlanningSpaceForNewRes_Guid;
                }
                //else
                //newResourceUsed[Guids.ResourceUsedReference.PlanningSpace_Guid].Value = Project(owner)[Dinamika.Guids.ProjectManagementReference.TypeProjectElement.PlanningSpace_Guid].Value;
                /*
                                //в библиотеке TF - isActual
                                newResourceUsed[Guids.ResourceUsedReference.FactValue_Guid].Value = usedResource[Guids.ResourceUsedReference.FactValue_Guid].Value;

                                //ResourceGroupLin
                                newResourceUsed.SetLinkedObject(Guids.ResourceUsedReference.Links.ResourceGroup_Nto1_Guid, usedResource.GetResourceGroupLink());

                                newResourceUsed[Guids.ResourceUsedReference.Workload_Guid].Value = usedResource[Guids.ResourceUsedReference.Workload_Guid].Value;
                                */
                newResourceUsed.EndChanges();

                // newResourceUsed.Reload();

                result.Add(newResourceUsed);
            }


            //подключаем используемые ресурсы к элемент укрумения
            if (МожноЛиРедактироватьОбъект(owner) == false) return false;

            owner.BeginChanges();

            foreach (var newResource in result)
            {
                // if (newResource.Changing)
                //     newResource.EndChanges();
                AddLinkedUsedResources(owner, newResource);
                //owner.AddLinkedObject(Guids.ProjectManagementReference.TypeProjectElement.Links.ProjectElementResLinks_1toN_Guid, newResource);
            }

            owner.EndChanges();

            owner.Unlock();

            //пересчитываем суммарные трудозатраты
            RecalcResourcesWorkLoad(owner);

            return true;
        }

        /// <summary>
        /// Проект элементаПроекта
        /// </summary>
        /// <param name="projectElement"></param>
        /// <returns></returns>
        public ReferenceObject Project(ReferenceObject projectElement)
        {
            if (projectElement == null)
            {
                return null;
            }

            if (projectElement.Class.IsInherit(TFlex.DOCs.References.ProjectManagement.ProjectManagementTypes.Keys.Project))
            {
                return projectElement.GetActual();
            }
            return Project(projectElement.Parent);
        }
        //cde1942c-ffc6-4654-9a2a-8f0b3a0a8e14

        /// <summary>
        /// Можно ли макросом изменить процент выполнения
        /// </summary>
        /// <param name="projectElement"></param>
        /// <returns></returns>
        public bool IsProgressReadOnly(ReferenceObject projectElement)
        {
            return projectElement[Guids.ProjectManagementReference.TypeProjectElement.DeterminationPercent_Guid].GetInt16() != 0;
        }

        /// <summary>
        /// Рекурсивный метод для пересчёта процентов с низу в верх
        /// </summary>
        /// <param name="projectElement">Элемент в котором требуется пересчёт процента</param>
        /// <returns></returns>
        public double ПересчитатьПроцентыЭлементаПроекта_СнизуДоВыбранного(ReferenceObject projectElement)
        {
            if (projectElement == null) return 0;

            double resultProgress = 0.0;

            double NumbHoursResource = 0.0;

            double Numerator = 0.0;

            if (!projectElement.HasChildren) return ПроцентВыполненияЭлементаПроекта(projectElement);

            foreach (var pe in projectElement.Children.Where(obj => obj.Class.IsInherit(Guids.ProjectManagementReference.TypeProjectElement_Guid)))
            {
                RecalcResourcesWorkLoad(pe);
                NumbHoursResource += ((TFlex.DOCs.References.ProjectManagement.ProjectElement)pe).SummaryPlannedResourcesWorkLoad;
                Numerator += ПересчитатьПроцентыЭлементаПроекта_СнизуДоВыбранного(pe) * ((TFlex.DOCs.References.ProjectManagement.ProjectElement)pe).SummaryPlannedResourcesWorkLoad;
            }

            if (NumbHoursResource > 0)
            {
                RecalcResourcesWorkLoad(projectElement);
                resultProgress = /*Math.Round(*/Numerator / ((TFlex.DOCs.References.ProjectManagement.ProjectElement)projectElement).SummaryPlannedResourcesWorkLoad/*, 5, MidpointRounding.AwayFromZero)*/;
            }

            if (МожноЛиРедактироватьОбъект(projectElement) == false) return resultProgress;

            ReferenceObject project = Project(projectElement);
            if (IsProgressReadOnly(project))
            {
                if (!project.Changing)
                    project.BeginChanges();
                project[Guids.ProjectManagementReference.TypeProjectElement.DeterminationPercent_Guid].Value = 0;
                project.EndChanges();
            }

            projectElement.BeginChanges();
            projectElement[Guids.ProjectManagementReference.TypeProjectElement.Percent_Guid].Value = resultProgress;
            projectElement.EndChanges();
            projectElement.Unlock();

            RecalcResourcesWorkLoad(projectElement);

            return resultProgress;
        }


        /// <summary>
        ///  Проверить возможность редактирования используемых ресурсов элемента проекта
        /// true - можно редактировать
        /// </summary>
        public bool МожноЛиРедактироватьИспользуемыеРесурсыЭлементаПроекта(ReferenceObject projectElement)
        {
            if (projectElement == null) return false;

            //Получение используемых ресурсов элемента укрупнения
            List<ReferenceObject> projectElmnt_UsedRes = new List<ReferenceObject>();


            projectElmnt_UsedRes.AddRange(GetLinkedUsedResources(projectElement)//projectElement.GetObjects(Dinamika.Guids.ProjectManagementReference.TypeProjectElement.Links.ProjectElementResLinks_1toN_Guid)
                       .Where(r => r.Class.IsInherit(ResourceUsed.Classes.Find
                      (Guids.ResourceUsedReference.TypeNonExpendableResourcesUsedReferenceObject_Guid))).ToList()
                       );

            List<ReferenceObject> BlockedUsedResources = new List<ReferenceObject>();

            BlockedUsedResources = projectElmnt_UsedRes.Where(
                res =>
                {
                    res.Unlock();
                    if (!res.CanEdit) return true;
                    else return false;
                }).ToList();

            if (BlockedUsedResources.Count() != 0)
            {
                return false;
            }

            return true;
        }
        /// <summary>
        /// Guid параметра элемента проекта Дата фактического начала
        /// </summary>
        private Guid ActualStartDate_guid
        {
            get
            {
                return Guids.ProjectManagementReference.TypeProjectElement.FactBeginDate_Guid;
            }
        }
        /// <summary>
        ///  "Дата фактичекого начала"
        /// </summary>
        /// <param name="projectElement"></param>
        /// <returns></returns>
        public TFlex.DOCs.Model.Parameters.DateTimeParameter ActualStartDate(ReferenceObject projectElement)
        {
            return (TFlex.DOCs.Model.Parameters.DateTimeParameter)projectElement[Guids.ProjectManagementReference.TypeProjectElement.FactBeginDate_Guid];
        }
        /// <summary>
        ///  "Дата фактичекого окончания"
        /// </summary>
        /// <param name="projectElement"></param>
        /// <returns></returns>
        public TFlex.DOCs.Model.Parameters.DateTimeParameter ActualEndDate(ReferenceObject projectElement)
        {
            return (TFlex.DOCs.Model.Parameters.DateTimeParameter)projectElement[Guids.ProjectManagementReference.TypeProjectElement.FactEndDate_Guid];
        }
        /// <summary>
        /// Получить статус работы
        /// </summary>
        /// <param name="projectElement">Элемент проекта</param>
        /// <returns>статус работы: Выполнено, В работе, Просроченная, Не выполнена. Eсли projectElement == string.Empty;</returns>
        public string СтатусРаботыЭлементаПроекта(ReferenceObject projectElement)
        {
            if (projectElement == null) return string.Empty;

            if (ПроцентВыполненияЭлементаПроекта(projectElement) == 1) return "Выполнено";//Выполнено

            if (projectElement[Guids.ProjectManagementReference.TypeProjectElement.EndDate_Guid].GetDateTime().Date.CompareTo(DateTime.Now.Date) >= 0)//плановое окончание позже или сегодняшнего числа (ещё не наступило)
            {

                if ((!ActualStartDate(projectElement).IsEmpty && ActualEndDate(projectElement).IsEmpty) ||
                    (ПроцентВыполненияЭлементаПроекта(projectElement) > 0 && ПроцентВыполненияЭлементаПроекта(projectElement) < 1)) // 0 < % < 1
                    return "В работе";//В работе

            }
            else //плановое окончание раньше сегодняшнего числа (уже прошло)
            {
                if (!projectElement[Guids.ProjectManagementReference.TypeProjectElement.EndDate_Guid].IsEmpty &&
                    ПроцентВыполненияЭлементаПроекта(projectElement) != 1) //процент выполнения не 100%
                    return "Просроченная";//Просроченная

                if (/*!work[factStart].IsEmpty && */ActualEndDate(projectElement).IsEmpty && //есть факт. старт и нет факт. окончания
                    ПроцентВыполненияЭлементаПроекта(projectElement) != 1) //процент выполнения не 100%
                    return "Не выполнена";//Не выполнена
            }


            return string.Empty;
        }
        /// <summary>
        /// Получить всех  пользователей , входящих в  организационную структуру
        /// </summary>
        /// <returns></returns>
        public List<User> ПолучитьВсехПользователейОрганизационнойСтруктуры()
        {
            var ДинамикаОргСтруктура = UserReference.Find(Guids.UserReference.Objects.Enterprise_Guid) as UserReferenceObject;
            return ДинамикаОргСтруктура.GetAllInternalUsers();
        }
        /// <summary>
        /// Получить всех  пользователей и групп, входящих в  организационную структуру
        /// </summary>
        /// <returns></returns>
        public List<UserReferenceObject> ПолучитьВсехПользоваталей_и_Групп_организационнойСтруктуры()
        {
            var ДинамикаОргСтруктура = UserReference.Find(Guids.UserReference.Objects.Enterprise_Guid) as UserReferenceObject;
            return ДинамикаОргСтруктура.GetAllInternalUsersAndGroups();
        }
        /// <summary>
        /// Получить ответственного проекта
        /// </summary>
        /// <param name="projectElement">Элемент проекта</param>
        /// <returns>null  - если projectElement или если нет ответственного</returns>
        public User ОтветвенныйПроектаИлиКубика1С_ЭлементаПроекта(ReferenceObject projectElement)
        {
            if (projectElement == null) return null;

            User Responsible = Project(projectElement).GetObject(Guids.ProjectManagementReference.TypeProjectElement.Links.Responsible_Nto1_Guid) as User;

            if (Responsible == null)
            {
                List<ReferenceObject> allchild = Project(projectElement).Children.RecursiveLoad();

                ReferenceObject Work1C = allchild.Where(h => h.Class.Name == "Затраты проекта 1С").FirstOrDefault();

                if (Work1C == null)
                {
                    return null;
                }
                else
                {
                    Responsible = Work1C.GetObject(Guids.ProjectManagementReference.TypeProjectElement.Links.Responsible_Nto1_Guid) as User;

                    if (Responsible == null) { return null; }
                    else
                    {
                        return Responsible;
                    }
                }
            }
            else
            {
                return Responsible;
            }
        }

        /// <summary>
        /// Получить список  всех подразделений(ресурсы укрупняются до подразделений) используемых ресурсов элемента проекта.
        /// </summary>
        /// <param name="projectElement">Элемент проекта, с которого нужно получить подразделения используемых ресурсов</param>
        /// <returns>null если projectElement == null</returns>
        public List<ReferenceObject> ПодразделенияТипаРесурс_ИспользуемыхРесурсовЭлементаПроекта(ReferenceObject projectElement)
        {
            List<ReferenceObject> result = new List<ReferenceObject>();

            if (projectElement == null) return result;

            //Используемые Ресурсы несинхронизированных  работ
            List<ReferenceObject> resourcesUsed = new List<ReferenceObject>();

            //Получаем список используемых ресурсов работы
            resourcesUsed = (GetLinkedUsedResources(projectElement))
                               //projectElement.GetObjects(Guids.ProjectManagementReference.TypeProjectElement.Links.ProjectElementResLinks_1toN_Guid))
                               /*.Where(/*h => h.IsPlanned())
                                */.ToList();

            //ReferenceObject Contractors = ProjectResourcesReference.Find(Guids.ProjectResourcesReference.Objects.Contractors_Guid);
            //ReferenceObject OKB = ProjectResourcesReference.Find(Guids.ProjectResourcesReference.Objects.OKB_Guid);
            //ReferenceObject ProductionSubdivision = ProjectResourcesReference.Find(Guids.ProjectResourcesReference.Objects.ProductionSubdivision_Guid);
            //ReferenceObject SMTS = ProjectResourcesReference.Find(Guids.ProjectResourcesReference.Objects.SMTS_Guid);

            foreach (ReferenceObject resourceUsed in resourcesUsed)
                result.Add(ПолучитьПодразделениеТипаРесурс_ИспользуемогоРесурса(resourceUsed));

            return result;
        }

        /// <summary>
        /// Поиск просроченных дочерних работ (и их дочерних) суммарного элемента проекта
        /// </summary>
        /// <param name="projectElement"></param>
        /// <returns></returns>
        public List<ReferenceObject> ПросроченныеРаботыЭлементаПроекта(ReferenceObject projectElement)
        {
            List<ReferenceObject> result = new List<ReferenceObject>();

            if (projectElement == null) return result;

            if (!projectElement.HasChildren) return result;

            List<ReferenceObject> allChildrenProjectElemet = projectElement.Children.RecursiveLoad();

            foreach (var work in allChildrenProjectElemet.Where(obj => obj.Class.IsInherit(Guids.ProjectManagementReference.TypeProjectElement_Guid)))
            {

                switch (СтатусРаботыЭлементаПроекта(work))
                {
                    case "Просроченная":
                        {
                            result.Add(work);
                            continue;
                        }
                    default: { continue; }
                }

            }
            if (result.Count == 0) return result;
            return result;
        }
        /// <summary>
        /// Группа используемого ресурса
        /// </summary>
        /// <param name="projectElement"></param>
        /// <returns></returns>
        public ReferenceObject GetResourceGroupLink(ReferenceObject projectElement)
        {

            return projectElement.GetObject(Guids.ResourceUsedReference.Links.ResourceGroup_Nto1_Guid);

            /*  set
              {
                  base.SetLinkedObject(ResourcesUsedReferenceObject.RelationKeys.ResourceGroupLink, value);
              }*/
        }
        /// <summary>
        /// Получить связанный ресурс
        /// </summary>
        /// <param usedResource=""></param>
        /// <returns></returns>
        public ReferenceObject GetResourcesLink(ReferenceObject usedResource)
        {
            return usedResource.GetObject(Guids.ResourceUsedReference.Links.ResourcesLink_Guid);
        }


        /// <summary>
        /// Получить используемые ресурсы элемента проекта
        /// </summary>
        /// <param name="projectElement">Элемент проекта</param>
        /// <returns></returns>
        public List<ReferenceObject> GetLinkedUsedResources(ReferenceObject projectElement)
        {
            return projectElement.GetObjects(Guids.ProjectManagementReference.TypeProjectElement.Links.ProjectElementResLinks_1toN_Guid);
        }

        /// <summary>
        ///  Добавить используемые ресурсы в элемент проекта
        /// </summary>
        /// <param name="projectElement">Куда цепляем</param>
        /// <param name="usedResource">Что цепляем</param>
        public void AddLinkedUsedResources(ReferenceObject projectElement, ReferenceObject usedResource)
        {
            projectElement.AddLinkedObject(Guids.ProjectManagementReference.TypeProjectElement.Links.ProjectElementResLinks_1toN_Guid, usedResource);
        }

        /// <summary>
        /// Возвращает ресурс подразделения(ресурс укрупняется до подразделения) используемого ресурса
        /// </summary>
        /// <param name="resource">Используемый ресурс или ресурс</param>
        /// <returns>если resource == null return null, </returns>
        public ReferenceObject ПолучитьПодразделениеТипаРесурс_ИспользуемогоРесурса(ReferenceObject resource)
        {
            if (resource == null) return null;

            ReferenceObject NonExpendableResource = null;

            ClassObject TypeNonExpendableResourceUsed = GetClassObject(ResourceUsed, Guids.ResourceUsedReference.TypeNonExpendableResourcesUsedReferenceObject_Guid);

            //Проверка на тип
            ReferenceObject NonExpendableResourceUsed = resource.Class.IsInherit(TypeNonExpendableResourceUsed) ? resource : null;

            if (NonExpendableResourceUsed == null)
            {
                //Приводим к типу нерасходываемый ресурс справочника Ресурсы
                NonExpendableResource = resource;
            }
            else//Получаем Ресурс из справочника Ресурсы
            {
                //Получаем Ресурс из справочника Ресурсы
                ReferenceObject GroupUserResource = GetResourceGroupLink(NonExpendableResourceUsed);

                NonExpendableResource = GetResourceGroupLink(NonExpendableResourceUsed);

                // if (!NonExpendableResource.IsGroup)
                if (GroupUserResource != null)
                    NonExpendableResource = GroupUserResource;

                if (NonExpendableResource == null)
                    return null;

                //Интересуют только плановые
                //if (NonExpendableResourceUsed.IsPlanned() == false) return "";

                //Проверка ресурса на тип
                switch (NonExpendableResource.Class.Name)
                {
                    case "Оборудование и оснастка":
                        { return null; }
                    case "Оснащение":
                        { return null; }
                    case "Комплектующие":
                        { return null; }
                    case "Ресурсы материалов":
                        { return null; }
                }
            }

            // NonExpendableResource
            var СвязанныйОбъектГиП = НайтиСвязанныйОбъектГиП_по_ресурсу(NonExpendableResource);

            ClassObject TypeUserGroup = GetClassObject(UserReference, Guids.UserReference.TypeUserGroup_Guid);

            if (СвязанныйОбъектГиП.Class.IsInherit(TypeUserGroup))
            {
                return СвязанныйОбъектГиП;
            }
            else
            {
                return ПолучитьПодразделениеТипаРесурс_ИспользуемогоРесурса(
                    НайтиСвязанныйРесурс_по_объектуГИП(СвязанныйОбъектГиП as UserReferenceObject).Parents.FirstOrDefault(p => Входит_в_оргСтруктуру(p as UserReferenceObject)).Parent);
            }
        }

        /// <summary>
        /// Получить все элементы проекта типа ЗатратыПроекта1С которые не являются отложенной версией и имеют отпределенный код1С по коду1С
        /// </summary>
        /// <returns> if (projectManagementReference == null) return null;</returns>
        public List<ReferenceObject> ВсеКубики1С_УП()
        {
            ClassObject TypeProjectCosts1C = ProjectManagementReference.Classes.Find(Guids.ProjectManagementReference.TypeProjectElement.TypeProjectCosts1C_Guid);

            Filter filter = new Filter(ProjectManagementReferenceInfo);

            ReferenceObjectTerm ИдентификаторАктуальногоЭлемента = new ReferenceObjectTerm(filter.Terms);
            // устанавливаем параметр
            ИдентификаторАктуальногоЭлемента.Path.AddParameter(ProjectManagementReference.ParameterGroup.OneToOneParameters.Find(Guids.ProjectManagementReference.TypeProjectElement.ActualElementId_Guid));
            // устанавливаем оператор сравнения
            ИдентификаторАктуальногоЭлемента.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            ИдентификаторАктуальногоЭлемента.Value = Guid.Empty;

            // Условие: тип элемента проекта порождён от типа ЗатратыПроекта1С
            ReferenceObjectTerm ТипЗатратыПроекта1С = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
            // устанавливаем параметр
            ТипЗатратыПроекта1С.Path.AddParameter(ProjectManagementReference.ParameterGroup[SystemParameterType.Class]);
            // устанавливаем оператор сравнения
            ТипЗатратыПроекта1С.Operator = ComparisonOperator.IsInheritFrom;
            // устанавливаем значение для оператора сравнения
            ТипЗатратыПроекта1С.Value = TypeProjectCosts1C;


            return ProjectManagementReference.Find(filter).ToList();
        }

        /// <summary>
        /// Найти проект в базе проектов по коду1С
        /// </summary>
        /// <returns> if (projectManagementReference == null) return null;</returns>
        public List<ReferenceObject> НайтиПроекты_в_БП_По_Коду1С(string code1C)
        {
            ReferenceInfo BaseProjectRefInfo = Connection.ReferenceCatalog.Find(Guids.BaseProject.Ref_Guid);
            Reference BaseProject = BaseProjectRefInfo.CreateReference();
            List<ReferenceObject> projects = BaseProject.Find(GetParameterInfo(BaseProjectRefInfo, Guids.BaseProject.Code1C_Guid), code1C);
            return projects;
        }

        /// <summary>
        /// Найти проект в базе проектов по коду1С
        /// </summary>
        /// <returns> if (projectManagementReference == null) return null;</returns>
        public List<ReferenceObject> НайтиПроекты_в_УП_По_Коду1С(string code1C)
        {
            List<ReferenceObject> projects = ProjectManagementReference.Find(GetParameterInfo(ProjectManagementReferenceInfo, Guids.ProjectManagementReference.TypeProjectElement.Code1C_Guid), code1C);
            return projects;
        }

        /// <summary>
        /// Возможность редактирования объекта, если можно true, иначе   false
        /// </summary>
        public bool МожноЛиРедактироватьОбъект(ReferenceObject projectElement)
        {
            if (projectElement == null) return false;

            projectElement.Unlock();
            if (projectElement.CanEdit)
            {
                try
                {
                    projectElement.BeginChanges();
                    projectElement.EndChanges();

                }
                catch (ObjectLockErrorException)
                {
                    return false;
                }
                finally
                {
                    projectElement.Unlock();
                }

                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Получить детализированные работы
        /// </summary>
        /// <param name="projectElement"></param>
        /// <returns>Список детализированных работа</returns>
        public List<ReferenceObject> ДетализацияЭлементаПроекта(ReferenceObject projectElement)
        {
            if (projectElement == null) return null;

            List<ReferenceObject> DetailingWorks = new List<ReferenceObject>();

            ReferenceObject DetailingWork = null;

            var Зависимости = СписокЗависимостейЭлементаПроектаТипаСинхронизация(projectElement, true).ToList();

            for (int i = 0; i < Зависимости.Count; i++)
            {
                Guid DetailingWorkGuid = (Guid)Зависимости[i][Guids.DependenciesReference.Work2_Guid].Value;//Получаем гуид Детализированной работы2

                Guid УкрупнённаяРаботаGuid = (Guid)Зависимости[i][Guids.DependenciesReference.Work1_Guid].Value;//Получаем гуид Укрупнённой работы1

                if (УкрупнённаяРаботаGuid != (Guid)projectElement.SystemFields.Guid) continue;

                DetailingWork = ProjectManagementReference.Find(DetailingWorkGuid);

                if (DetailingWork == null || DetailingWork.IsInRecycleBin) continue;

                DetailingWorks.Add(DetailingWork);
            }

            return DetailingWorks;
        }


        /// <summary>
        /// Плановый ресурс
        /// </summary>
        /// <param name="ResUsed"></param>
        /// <returns></returns>
        public bool IsPlanned(ReferenceObject ResUsed)
        {
            return !ResUsed[Guids.ResourceUsedReference.FactValue_Guid].GetBoolean();
        }
        /// <summary>
        /// Процент выполнения элемента проекта
        /// </summary>
        /// <param name="projectElement">Элемент проекта</param>
        /// <returns></returns>
        public double ПроцентВыполненияЭлементаПроекта(ReferenceObject projectElement)
        {
            return projectElement[Guids.ProjectManagementReference.TypeProjectElement.Percent_Guid].GetDouble();
        }


        /// <summary>
        ///  Удаление  использумых ресурсов c элемента проекта
        /// </summary>
        /// <param name="projectElement">"Элемент проекта</param>
        /// <param name="onlyPlanenned">Удалить только плановые ресурсы</param>
        /// <returns></returns>
        public bool УдалитьИспользуемыеРесурсыЭлементаПроекта(ReferenceObject projectElement, bool onlyPlanenned = false)
        {
            if (projectElement == null) return false;

            List<ReferenceObject> projectElmnt_UsedRes = new List<ReferenceObject>();
            // projectElmnt_UsedRes.AddRange(GetLinkedUsedResources(projectElement).Where(r => r.Class.IsInherit(Dinamika.ExtensionDinamika.ResourceUsed.Classes.Find
            projectElmnt_UsedRes.AddRange(GetLinkedUsedResources(projectElement)
                                   .Where(r => r.Class.IsInherit(ResourceUsed.Classes.Find
                                  (Guids.ResourceUsedReference.TypeNonExpendableResourcesUsedReferenceObject_Guid))).ToList()
                                   );


            if (МожноЛиРедактироватьОбъект(projectElement) == false) return false;

            projectElement.BeginChanges();

            //Удаление используемых ресурсов укрупнения
            foreach (var resourcesUsed in projectElmnt_UsedRes)
            {
                if (IsPlanned(resourcesUsed) && onlyPlanenned)
                    resourcesUsed.Delete();
                else
                    resourcesUsed.Delete();
            }

            projectElement.EndChanges();

            projectElement.Unlock();

            return true;
        }


        /// <summary>
        /// Удаление  использумых ресурсов с суммарной работы
        /// </summary>
        /// <param name="projectElement">"Элемент проекта</param>
        /// <param name="onlyPlanenned">Удалить только плановые ресурсы</param>
        /// <returns></returns>
        public void УдалениеИспользуемыхРесурсовСуммарногоЭлементаПроекта(ReferenceObject projectElement, bool onlyPlanenned = false)
        {
            if (projectElement == null || !projectElement.HasChildren) return;

            if (!projectElement.HasChildren) return;

            List<ReferenceObject> projectElmnt_UsedRes = new List<ReferenceObject>();

            projectElmnt_UsedRes.AddRange(GetLinkedUsedResources(projectElement)
                                               .Where(r => r.Class.IsInherit(ResourceUsed.Classes.Find
                                              (Guids.ResourceUsedReference.TypeNonExpendableResourcesUsedReferenceObject_Guid))).ToList()
                                               );

            if (МожноЛиРедактироватьОбъект(projectElement) == false) return;

            projectElement.BeginChanges();

            //Удаление используемых ресурсов укрупнения
            foreach (var resourcesUsed in projectElmnt_UsedRes)
            {
                if (IsPlanned(resourcesUsed) && onlyPlanenned)
                    resourcesUsed.Delete();
                else
                    resourcesUsed.Delete();

            }

            projectElement.EndChanges();
            projectElement.Unlock();

        }

        /// <summary>
        /// Пересчить Суммарные значения ресурсов
        /// </summary>
        /// <param name="pe"></param>
        /// <param name="flag"></param>
        /// <param name="flag2"></param>
        public void RecalcResourcesWorkLoad(ReferenceObject pe, bool flag = false, bool flag2 = true)
        {
            var PE = pe as TFlex.DOCs.References.ProjectManagement.ProjectElement;

            if (PE == null) return;
            PE.RecalcResourcesWorkLoad(flag, flag2);
        }

        #endregion


        #region расширяющие методы Группы и пользователи, Ресурсы


        /// <summary>
        /// ПорученияПользователяСрокВыполненияКоторыхЗаканчиваетсяСегодня
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public List<ReferenceObject> ПорученияПользователяСрокВыполненияКоторыхЗаканчиваетсяСегодня(ReferenceObject user)
        {
            List<ReferenceObject> findedObjects = new List<ReferenceObject>();

            TFlex.DOCs.Model.Stages.Stage stage_В_работе = TFlex.DOCs.Model.Stages.Stage.Find(Connection, "В работе_Контроль поручений");

            DateTime today = DateTime.Now.Date;

            //Формирование фильтра
            using (Filter filter = new Filter(ErrandControlReferenceInfo))
            {
                //стадия == В работе_Контроль поручений
                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(SystemParameterType.Stage),
                                     ComparisonOperator.Equal, stage_В_работе);

                //исполнитель текущий пользователь
                filter.Terms.AddTerm(Guids.ErrandControl.Links.Performer_Guid.ToString(), ComparisonOperator.Equal, user);

                //Плановая дата выполнения поручения >= заданного
                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.EndDate_Guid),
                                     ComparisonOperator.GreaterThanOrEqual, today);

                //Плановая дата выполнения поручения < завтра
                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.EndDate_Guid),
                                     ComparisonOperator.LessThan, today.AddDays(1));


                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.PerformerResult_Guid),
                                     ComparisonOperator.NotEqual, "Выполнено");

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.PerformerResult_Guid),
                     ComparisonOperator.NotEqual, "Коррекция");

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.ValidationResult_Guid),
                                     ComparisonOperator.NotEqual, "Аннулировано");

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.ValidationResult_Guid),
                     ComparisonOperator.NotEqual, "Выполнено");

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.ValidationResult_Guid),
        ComparisonOperator.NotEqual, "Коррекция");

                //Применяем фильтр
                findedObjects = ErrandControl.Find(filter);
            }

            return findedObjects;
        }
 

        /// <summary>
        /// НеПринятыеПользователемПоручения
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public List<ReferenceObject> НеПринятыеПользователемПоручения(ReferenceObject user)
        {
            List<ReferenceObject> findedObjects = new List<ReferenceObject>();

            TFlex.DOCs.Model.Stages.Stage stage_В_работе = TFlex.DOCs.Model.Stages.Stage.Find(Connection, "В работе_Контроль поручений");

            //Формирование фильтра
            using (Filter filter = new Filter(ErrandControlReferenceInfo))
            {
                //стадия == В работе_Контроль поручений
                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(SystemParameterType.Stage),
                                     ComparisonOperator.Equal, stage_В_работе);

                //исполнитель текущий пользователь
                filter.Terms.AddTerm(Guids.ErrandControl.Links.Performer_Guid.ToString(), ComparisonOperator.Equal, user);

                //Плановая дата выполнения поручения > заданного
                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.PerformerResult_Guid),
                                     ComparisonOperator.Equal, "Не принято в работу");

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.ValidationResult_Guid),
                                     ComparisonOperator.NotEqual, "Аннулировано");

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.ValidationResult_Guid),
                     ComparisonOperator.NotEqual, "Выполнено");

                //Применяем фильтр
                findedObjects = ErrandControl.Find(filter);
            }

            return findedObjects;
        }

        /// <summary>
        /// ПросроченныеПользователемПоручения
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public List<ReferenceObject> ПросроченныеПользователемПоручения(ReferenceObject user)
        {
            List<ReferenceObject> findedObjects = new List<ReferenceObject>();

            TFlex.DOCs.Model.Stages.Stage stage_В_работе = TFlex.DOCs.Model.Stages.Stage.Find(Connection, "В работе_Контроль поручений");

            DateTime today = DateTime.Now.Date;

            //Формирование фильтра
            using (Filter filter = new Filter(ErrandControlReferenceInfo))
            {
                //стадия == В работе_Контроль поручений
                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(SystemParameterType.Stage),
                                     ComparisonOperator.Equal, stage_В_работе);

                //исполнитель текущий пользователь
                filter.Terms.AddTerm(Guids.ErrandControl.Links.Performer_Guid.ToString(), ComparisonOperator.Equal, user);

                //Плановая дата выполнения поручения < Текущей даты
                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.EndDate_Guid),
                                     ComparisonOperator.LessThan, today);

                //ValidationResult_Guid
                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.PerformerResult_Guid),
                         ComparisonOperator.Equal, "Не выполнено");

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.ValidationResult_Guid),
                     ComparisonOperator.NotEqual, "Выполнено");

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.ValidationResult_Guid),
                     ComparisonOperator.NotEqual, "Аннулировано");


                //Применяем фильтр
                findedObjects = ErrandControl.Find(filter);
            }

            return findedObjects;
        }

        /// <summary>
        /// ТребуетсяОтПользователяРуководителяКоррекцияПоручения
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public List<ReferenceObject> ТребуетсяОтПользователяРуководителяКоррекцияПоручения(ReferenceObject user)
        {
            List<ReferenceObject> findedObjects = new List<ReferenceObject>();

            TFlex.DOCs.Model.Stages.Stage stage_В_работе = TFlex.DOCs.Model.Stages.Stage.Find(Connection, "В работе_Контроль поручений");

            DateTime today = DateTime.Now.Date;

            //Формирование фильтра
            using (Filter filter = new Filter(ErrandControlReferenceInfo))
            {
                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(SystemParameterType.Stage),
                           ComparisonOperator.Equal, stage_В_работе);

                filter.Terms.AddTerm(Guids.ErrandControl.Links.Director_Guid.ToString(), ComparisonOperator.Equal, user);

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.PerformerResult_Guid),
                         ComparisonOperator.Equal, "Требуется коррекция");

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.ValidationResult_Guid),
        ComparisonOperator.NotEqual, "Выполнено");

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.ValidationResult_Guid),
                     ComparisonOperator.NotEqual, "Аннулировано");

                //Применяем фильтр
                findedObjects = ErrandControl.Find(filter);
            }

            return findedObjects;
        }

        /// <summary>
        /// ВыданныеПользователемПорученияСрокКоторыхИстёк
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public List<ReferenceObject> ВыданныеПользователемПорученияСрокКоторыхИстёк(ReferenceObject user)
        {
            List<ReferenceObject> findedObjects = new List<ReferenceObject>();

            TFlex.DOCs.Model.Stages.Stage stage_В_работе = TFlex.DOCs.Model.Stages.Stage.Find(Connection, "В работе_Контроль поручений");

            DateTime today = DateTime.Now.Date;

            //Формирование фильтра
            using (Filter filter = new Filter(ErrandControlReferenceInfo))
            {
                //стадия в работе
                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(SystemParameterType.Stage),
                           ComparisonOperator.Equal, stage_В_работе);

                filter.Terms.AddTerm(Guids.ErrandControl.Links.Director_Guid.ToString(), ComparisonOperator.Equal, user);

                //Плановая дата выполнения поручения < Текущей даты
                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.EndDate_Guid),
                                     ComparisonOperator.LessThan, today);

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.PerformerResult_Guid),
         ComparisonOperator.Equal, "Не выполнено");

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.ValidationResult_Guid),
        ComparisonOperator.NotEqual, "Выполнено");

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.ValidationResult_Guid),
                     ComparisonOperator.NotEqual, "Аннулировано");

                //Применяем фильтр
                findedObjects = ErrandControl.Find(filter);
            }

            return findedObjects;
        }

        /// <summary>
        /// ВыданныеПользователемПоручения
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public List<ReferenceObject> ВыданныеПользователемПоручения(ReferenceObject user)
        {
            List<ReferenceObject> findedObjects = new List<ReferenceObject>();

            TFlex.DOCs.Model.Stages.Stage stage_В_работе = TFlex.DOCs.Model.Stages.Stage.Find(Connection, "В работе_Контроль поручений");

            DateTime today = DateTime.Now.Date;

            //Формирование фильтра
            using (Filter filter = new Filter(ErrandControlReferenceInfo))
            {
                //стадия в работе
                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(SystemParameterType.Stage),
                           ComparisonOperator.Equal, stage_В_работе);

                filter.Terms.AddTerm(Guids.ErrandControl.Links.Director_Guid.ToString(), ComparisonOperator.Equal, user);

                //Плановая дата выполнения поручения < Текущей даты
                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.EndDate_Guid),
                                     ComparisonOperator.GreaterThanOrEqual, today);

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.PerformerResult_Guid),
         ComparisonOperator.Equal, "Не выполнено");

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.ValidationResult_Guid),
        ComparisonOperator.NotEqual, "Выполнено");

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.ValidationResult_Guid),
                     ComparisonOperator.NotEqual, "Аннулировано");

                //Применяем фильтр
                findedObjects = ErrandControl.Find(filter);
            }

            return findedObjects;
        }

        /// <summary>
        /// НеПрипятыеОтПользователяПоручения
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public List<ReferenceObject> НеПринятыеОтПользователяПоручения(ReferenceObject user)
        {
            List<ReferenceObject> findedObjects = new List<ReferenceObject>();

            TFlex.DOCs.Model.Stages.Stage stage_В_работе = TFlex.DOCs.Model.Stages.Stage.Find(Connection, "В работе_Контроль поручений");

            DateTime today = DateTime.Now.Date;

            //Формирование фильтра
            using (Filter filter = new Filter(ErrandControlReferenceInfo))
            {
                //стадия в работе
                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(SystemParameterType.Stage),
                           ComparisonOperator.Equal, stage_В_работе);

                filter.Terms.AddTerm(Guids.ErrandControl.Links.Director_Guid.ToString(), ComparisonOperator.Equal, user);

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.PerformerResult_Guid),
             ComparisonOperator.Equal, "Не принято в работу");

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.ValidationResult_Guid),
        ComparisonOperator.NotEqual, "Выполнено");

                filter.Terms.AddTerm(ErrandControl.ParameterGroup.OneToOneParameters.Find(Guids.ErrandControl.ValidationResult_Guid),
                     ComparisonOperator.NotEqual, "Аннулировано");


                //Применяем фильтр
                findedObjects = ErrandControl.Find(filter);
            }

            return findedObjects;
        }

        /// <summary>
        /// СлужебныеЗапискиНаПодпись
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public List<ReferenceObject> СлужебныеЗапискиПользователяНаПодпись(ReferenceObject user)
        {
            TFlex.DOCs.Model.Stages.Stage СЗСогласование = TFlex.DOCs.Model.Stages.Stage.Find(Connection, "СЗ. Согласование");
            TFlex.DOCs.Model.Stages.Stage СЗУтверждение = TFlex.DOCs.Model.Stages.Stage.Find(Connection, "СЗ. Утверждение");

            List<ReferenceObject> result = new List<ReferenceObject>();

            var dataOfficeNotes = OfficeNotes.Find(OfficeNotes.ParameterGroup[SystemParameterType.Stage], СЗСогласование);
            dataOfficeNotes.AddRange(OfficeNotes.Find(OfficeNotes.ParameterGroup[SystemParameterType.Stage], СЗУтверждение));

            foreach (var СЗ in dataOfficeNotes)
            {
                foreach (Signature sign in СЗ.Signatures)
                {
                    if (sign.UserName == user[Guids.UserReference.Name_Guid].GetString() && sign.SignatureDate == null)
                    {
                        result.Add(СЗ);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Найти подразделение по коду
        /// </summary>
        /// <param name="Code">Код, по которому нужно найти подразделение</param>
        /// <returns></returns>
        public ReferenceObject НайтиГруппуПользователейПоКоду(string Code)
        {
            if (string.IsNullOrEmpty(Code)) return null;

            ReferenceObject subdivision = null;
            using (Filter filter = new Filter(UserReferenceInfo))
            {
                filter.Terms.AddTerm(UserReference.ParameterGroup[SystemParameterType.Class], ComparisonOperator.IsInheritFrom, UserReference.Classes.GroupBaseType);
                filter.Terms.AddTerm(GetParameterInfo(UserReferenceInfo, Guids.UserReference.TypeUserGroup.Code_Guid), ComparisonOperator.Equal, Code);//Код
                subdivision = UserReference.Find(filter).FirstOrDefault();
            }
            return subdivision;
        }
        /// <summary>
        /// ПолучитьВсехПользователейГруппы
        /// </summary>
        /// <param name="UserReferenceObjects"></param>
        /// <returns></returns>
        public List<ReferenceObject> ПолучитьВсехПользователейГруппы(UserReferenceObject UserReferenceObjects)
        {
            return UserReferenceObjects.Children.RecursiveLoad().Where(obj => obj.Class.IsInherit(Connection.References.Users.Classes.UserBaseType)).ToList();
        }

        /// <summary>
        /// Найти пользователя по кодуФизЛица
        /// </summary>
        /// <param name="UserCode1C">КодФизЛица, по которому нужно найти сотрудника</param>
        /// <returns></returns>
        public ReferenceObject НайтиПользователяПоКодуФизЛица(string UserCode1C)
        {
            if (UserCode1C == null) return null;

            ReferenceObject user = null;
            using (Filter filter = new Filter(UserReferenceInfo))
            {
                filter.Terms.AddTerm(UserReference.ParameterGroup[SystemParameterType.Class], ComparisonOperator.IsInheritFrom, UserReference.Classes.UserBaseType);
                filter.Terms.AddTerm(GetParameterInfo(UserReferenceInfo, Guids.UserReference.TypeUser.TypeEmployer.UserCode1C_Guid), ComparisonOperator.Equal, UserCode1C);//КодФизЛица
                user = UserReference.Find(filter).FirstOrDefault();
            }
            return user;
        }
        /// <summary>
        /// Найти пользователей по должности
        /// </summary>
        public List<User> НайтиПользоватейПоДолжности(string post)
        {
            List<ComplexHierarchyLink> usersLinks = UserReference.Objects.RecursiveLoadHierarchyLinks();//OfType<User>();

            List<User> Result = new List<User>();
            foreach (var userLink in usersLinks)
            {
                if (userLink[Guids.UserReference.ParametersHierarchy_Post_Guid].GetString().StartsWith(post))
                {
                    User user = userLink.ChildObject as User;
                    if (user != null)
                        Result.Add(user);
                }
            }
            return Result;
        }
        /// <summary>
        /// Входит ли объект в гп ДОС
        /// </summary>
        /// <param name="userObject"></param>
        /// <returns></returns>
        public bool Входит_в_оргСтруктуру(UserReferenceObject userObject)
        {
            var ДинамикаОргСтруктура = UserReference.Find(Guids.UserReference.Objects.Enterprise_Guid) as UserReferenceObject;

            if (userObject == ДинамикаОргСтруктура) return true;

            List<ComplexHierarchyLink> hierarchyLinks = userObject.Parents.GetHierarchyLinks();

            if (hierarchyLinks.Any(hl => hl.ParentObject == ДинамикаОргСтруктура))
                return true;

            foreach (ComplexHierarchyLink hierarchyLink in hierarchyLinks)
            {
                if (Входит_в_оргСтруктуру(hierarchyLink.ParentObject as UserReferenceObject))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Возвращает список подключений пользователя к организационной структуре 
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public List<ComplexHierarchyLink> НайтиПодклПользователя_к_оргСтруктуре(User user)
        {
            List<ComplexHierarchyLink> HLinksActualPost = new List<ComplexHierarchyLink>();
            user.Parents.Reload();
            foreach (var Hlink in user.Parents.GetHierarchyLinks().Where(hl => Входит_в_оргСтруктуру(hl.ParentObject as UserReferenceObject)))
            {
                HLinksActualPost.Add(Hlink);
            }
            return HLinksActualPost;
        }
        /// <summary>
        /// Возвращает список подключений (у которых не пуст параметр Должность) пользователя к организационной структуре 
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public List<ComplexHierarchyLink> НайтиПодклПользователя_с_не_пустойДолжностью_к_оргСтруктуре(User user)
        {
            List<ComplexHierarchyLink> HLinksActualPost = new List<ComplexHierarchyLink>();
            user.Parents.Reload();
            foreach (var Hlink in user.Parents.GetHierarchyLinks().Where(hl => Входит_в_оргСтруктуру(hl.ParentObject as UserReferenceObject)))
            {
                if (!string.IsNullOrEmpty(Hlink[Guids.UserReference.ParametersHierarchy_Post_Guid].GetString()))
                {
                    HLinksActualPost.Add(Hlink);
                }
            }
            return HLinksActualPost;
        }

        /// <summary>
        /// Найти ресурс объекта спр. группы и пользователи
        /// </summary>
        /// <param name="userReferenceObject">объект, чей ресурс нужно найти</param>
        /// <returns></returns>
        public ReferenceObject НайтиСвязанныйРесурс_по_объектуГИП(UserReferenceObject userReferenceObject)
        {
            if (userReferenceObject == null) return null;
            ReferenceObject resource = null;

            //Ищем связанный ресурс с объектом user
            using (Filter filter = new Filter(ProjectResourcesReferenceInfo))
            {
                filter.Terms.AddTerm(Guids.ProjectResourcesReference.TypeNonExpendableResourcesReferenceObject.TypeHumanResources.Links.Human_UserRef_1to1_Guid.ToString(), ComparisonOperator.Equal, userReferenceObject);
                resource = ProjectResourcesReference.Find(filter).FirstOrDefault();
            }

            return resource;
        }
        /// <summary>
        /// Найти объекта спр. группы и пользователи по его ресурсу
        /// </summary>
        /// <param name="nonExpendableResourcesReferenceObject">объект, чей связанный объект в ГиП нужно найти</param>
        /// <returns></returns>
        public ReferenceObject НайтиСвязанныйОбъектГиП_по_ресурсу(ReferenceObject nonExpendableResourcesReferenceObject)
        {
            if (nonExpendableResourcesReferenceObject == null) return null;
            ReferenceObject userRefObj = null;

            //Ищем связанный ресурс с объектом user
            using (Filter filter = new Filter(UserReferenceInfo))
            {
                filter.Terms.AddTerm(Guids.ProjectResourcesReference.TypeNonExpendableResourcesReferenceObject.TypeHumanResources.Links.Human_UserRef_1to1_Guid.ToString(), ComparisonOperator.Equal, nonExpendableResourcesReferenceObject);
                userRefObj = UserReference.Find(filter).FirstOrDefault();
            }

            return userRefObj;
        }

        #endregion


        #region методы  спр.Файлы

        /// <summary>
        /// Добавить доступ на объект
        /// </summary>
        /// <param name="referenceObject">Объект на который нужно назначить доступ</param>
        /// <param name="userReferenceObject">Кому дать доступ</param>
        /// <param name="AccessName">Наименование доступа, например - "Просмотр"</param>
        /// <param name="useAccess">0 - Объект и его дочернии объекты, 1 - Только дочернии объекты, 2 - объект</param>
        public void ДобавитьДоступНаОбъект(ReferenceObject referenceObject, UserReferenceObject userReferenceObject, string AccessName, int useAccess = 0)
        {
            // Получаем все доступы
            List<AccessGroup> allAccessGroups = AccessGroup.GetGroups(Connection);

            AccessGroup Access = allAccessGroups.Find(ag => ag.Type.IsObject && ag.Name == AccessName);
            if (Access == null) return;

            AccessManager ДанныеОдоступахНаОбъект = AccessManager.GetReferenceObjectAccess(referenceObject);

            //Направление действия группы прав доступа 
            AccessDirection AccessDirection;

            if (useAccess == 0)
                AccessDirection = AccessDirection.Default;
            else if (useAccess == 1)
                AccessDirection = AccessDirection.Children;
            else if (useAccess == 2)
                AccessDirection = AccessDirection.Entity;
            else return;

            //Тип команд доступа 
            AccessCommandType AccessCommandType = AccessCommandType.Object;

            //Тип прав доступа в менеджере 
            AccessTypeID AccessTypeID = AccessTypeID.Object;

            // Снимаем явную установку доступа
            ДанныеОдоступахНаОбъект.SetInherit(false, false);

            ДанныеОдоступахНаОбъект.SetAccess(0, userReferenceObject, Access, Access, null, AccessCommandType
                , AccessTypeID, AccessDirection, AccessDirection);

            ДанныеОдоступахНаОбъект.Save();    // Сохраняем изменения
        }

        /// <summary>
        /// Добавить доступ на объект 
        /// </summary>
        /// <param name="referenceObject">Объект на который нужно назначить доступ</param>
        /// <param name="userReferenceObject">Кому дать доступ</param>
        /// <param name="AccessName">Наименование доступа, например - "Просмотр"</param>
        public void ДобавитьДоступНаОбъект(ReferenceObject referenceObject, UserReferenceObject userReferenceObject, string AccessName)
        {
            // Получаем все доступы
            List<AccessGroup> allAccessGroups = AccessGroup.GetGroups(Connection);

            AccessGroup Access = allAccessGroups.Find(ag => ag.Type.IsObject && ag.Name == AccessName);
            if (Access == null) return;

            AccessManager ДанныеОдоступахНаОбъект = AccessManager.GetReferenceObjectAccess(referenceObject);

            // Снимаем явную установку доступа
            ДанныеОдоступахНаОбъект.SetInherit(false, false);

            ДанныеОдоступахНаОбъект.SetAccess(userReferenceObject, Access); // Устанавливаем доступ

            ДанныеОдоступахНаОбъект.Save();    // Сохраняем изменения
        }

        /// <summary>
        /// Удалить  доступ с объекта
        /// </summary>
        /// <param name="referenceObject">Объект доступ с которого нужно удалить</param>
        /// <param name="userReferenceObject">Кому удалить доступ</param>
        /// <param name="AccessName">Наименование доступа, например - "Просмотр"</param>
        /// <param name="useAccess">0 - Объект и его дочернии объекты, 1 - Только дочернии объекты, 2 - объект</param>
        public void УдалитьДоступНаОбъект(ReferenceObject referenceObject, UserReferenceObject userReferenceObject, string AccessName, int useAccess = 0)
        {
            // Получаем все доступы
            List<AccessGroup> allAccessGroups = AccessGroup.GetGroups(Connection);

            AccessGroup Access = allAccessGroups.Find(ag => ag.Type.IsObject && ag.Name == AccessName);
            if (Access == null) return;


            AccessManager ДанныеОдоступахНаОбъект = AccessManager.GetReferenceObjectAccess(referenceObject, TFlex.DOCs.Common.AccessRightsLoadOptions.EditorMode);
            //AccessManager ДанныеОдоступахНаОбъект = AccessManager.GetReferenceObjectAccess(referenceObject);//При useAccess = 1. Доступ почему-то не удаляется

            AccessDirection AccessDirection;

            if (useAccess == 0)
                AccessDirection = AccessDirection.Default;
            else if (useAccess == 1)
                AccessDirection = AccessDirection.Children;
            else if (useAccess == 2)
                AccessDirection = AccessDirection.Entity;
            else return;

            var AccessType = Access.Type;

            // Снимаем явную установку доступа
            ДанныеОдоступахНаОбъект.SetInherit(false, false);

            ДанныеОдоступахНаОбъект.RemoveAccess(userReferenceObject, AccessType, AccessDirection);

            ДанныеОдоступахНаОбъект.Save();    // Сохраняем изменения
        }

        private ReferenceObject addInPersonalDataIdFolder(string UserCode1C, string idFolder)
        {
            var personalDataInfo = Connection.ReferenceCatalog.Find(Guids.DataEmployees.Ref_Guid);

            var personalData = personalDataInfo.CreateReference();

            var pData = personalData.Find(GetParameterInfo(personalDataInfo, Guids.DataEmployees.UserCode1C_Guid), UserCode1C).ToList().FirstOrDefault();

            if (pData == null) return pData;

            if (pData[Guids.DataEmployees.idFolder_Guid].GetString() != idFolder)
            {
                pData.Unlock();
                pData.BeginChanges();
                pData[Guids.DataEmployees.idFolder_Guid].Value = idFolder;
                pData.EndChanges();
            }
            return pData;
        }

        /// <summary>
        /// ПолучитьЛичнуюПапкуПользователя_Файлы
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public ReferenceObject ПолучитьЛичнуюПапкуПользователя_Файлы(User user)
        {
            var personalDataInfo = Connection.ReferenceCatalog.Find(Guids.DataEmployees.Ref_Guid);

            var personalData = personalDataInfo.CreateReference();

            var pData = personalData.Find(GetParameterInfo(personalDataInfo, Guids.DataEmployees.UserCode1C_Guid)
                , user[Guids.UserReference.TypeUser.TypeEmployer.UserCode1C_Guid].GetString()).ToList().FirstOrDefault();

            if (pData == null) return null;

            FileReference fileReference = Connection.ReferenceCatalog.Find(Guids.FileReference.Ref_Guid).CreateReference() as FileReference;

            if (string.IsNullOrEmpty(pData[Guids.DataEmployees.idFolder_Guid].GetString())) return null;

            var ЛичнаяПапка = fileReference.Find(fileReference.ParameterGroup[SystemParameterType.ObjectId]
                , pData[Guids.DataEmployees.idFolder_Guid].GetString()).ToList().FirstOrDefault();

            if (ЛичнаяПапка != null) return ЛичнаяПапка;
            else
            {
                string nameFolder = user.FullName;

                ЛичнаяПапка = fileReference.CreatePath("Личное\\" + nameFolder, null) as FolderObject; ;

                ReferenceObject publicПапка = fileReference.CreatePath("Личное\\" + nameFolder + "\\public", null) as FolderObject;

                Desktop.CheckIn(ЛичнаяПапка, "Создание и назначение доступа", false);

                Desktop.CheckIn(publicПапка, "Создание и назначение доступа", false);

                ДобавитьДоступНаОбъект(ЛичнаяПапка, user, "Просмотр");
                ДобавитьДоступНаОбъект(publicПапка, user, "Доступ к личной папке");
                ДобавитьДоступНаОбъект(publicПапка, null, "Просмотр");

                addInPersonalDataIdFolder(user[Guids.UserReference.TypeUser.TypeEmployer.UserCode1C_Guid].GetString(), ЛичнаяПапка.SystemFields.Id.ToString());

                return ЛичнаяПапка;
            }

        }


        /// <summary>
        /// ПолучитьПубличнуюЛичнуюПапку_Файлы
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public ReferenceObject ПолучитьПубличнуюЛичнуюПапку_Файлы(User user)
        {
            FileReference fileReference = Connection.ReferenceCatalog.Find(Guids.FileReference.Ref_Guid).CreateReference() as FileReference;

            var ЛичнаяПапка = ПолучитьЛичнуюПапкуПользователя_Файлы(user);

            string nameFolder = user.FullName;

            ReferenceObject publicПапка = fileReference.FindByRelativePath("Личное\\" + nameFolder + "\\public");

            if (publicПапка == null)
            {
                //System.Windows.Forms.MessageBox.Show("publicПапка == null");
                publicПапка = fileReference.CreatePath("Личное\\" + nameFolder + "\\public", null);
                Desktop.CheckIn(publicПапка, "Создание и назначение доступа", false);
                ДобавитьДоступНаОбъект(publicПапка, null, "Просмотр");
                ДобавитьДоступНаОбъект(publicПапка, user, "Доступ к личной папке", 0);
            }

            return publicПапка;
        }
        #endregion


        /// <summary>
        /// Проверка на пустоту
        /// </summary>
        /// <param name="sb"></param>
        /// <returns>True если пустой</returns>
        public bool IsStringBuilderNullOrEmpty(StringBuilder sb)
        {
            return (sb == null || sb.Length == 0);
        }


        /// <summary>
        /// Возвращает экземпляр справочника для работы с данными
        /// </summary>
        /// <param name="referenceInfo">Описание справочника</param>
        /// <returns></returns>
        public Reference GetReference(ReferenceInfo referenceInfo)
        {
            Reference reference = referenceInfo.CreateReference();

            return reference;
        }

        /// <summary>
        /// Является ли текущая дата рабочем днём по календарю Динамики
        /// </summary>
        /// <param name="currentDate"></param>
        /// <returns></returns>
        public bool РабочийДень(DateTime currentDate)// используется
        {
            //currentDate = new DateTime(2014,12,30);

            //CalendarReferenceObject calendar = new CalendarReference(servConnect).Find(new Guid("7f919a1c-a1e2-414b-9e3f-2d79f1a782c6")) as CalendarReferenceObject;//тестовый календарь
            CalendarReferenceObject calendar = new CalendarReference(Connection).Find(new Guid("561b03ce-feb0-4f4f-bf20-8c76e2f14bbe")) as CalendarReferenceObject; // Календарь АО ЦНТУ Динамика

            bool isWorking = true;

            DateTime date = currentDate;

            foreach (var workTime in calendar.WorkingTimes)
            {
                //MessageBox.Show( i.ToString() + ". " + workTime.ToString() +" - "+ workTime.ContainsDate(currentDate).ToString() + "\n" + workTime.WorkTimeIntervals.Count().ToString());
                if (workTime.ContainsDate(date) && workTime.WorkTimeIntervals.Count() == 0) isWorking = false;
            }

            if (!isWorking) return false;

            return true;
        }

        /// <summary>
        /// УдалениеВсехПодписей
        /// </summary>
        /// <param name="refObject"></param>
        public void УдалитьВсеПодписи(ReferenceObject refObject)
        {
            foreach (var sign in refObject.Signatures)
                sign.Delete();
        }
        /// <summary>
        /// удаляет все подписи указанного подписанта
        /// </summary>
        /// <param name="refObject"></param>
        /// <param name="Подписант"></param>
        /// <param name="ТипПодписи"></param>
        public void УдалитьПодписьПодписанта(ReferenceObject refObject, string Подписант, string ТипПодписи = null)
        {
            foreach (var sign in refObject.Signatures)
            {
                if (sign.UserName == Подписант && ТипПодписи == null)
                    sign.Delete();
                else
                    if (sign.UserName == Подписант && sign.SignatureObjectType.Name == ТипПодписи)
                    sign.Delete();
            }
        }
        /// <summary>
        /// Смена стадии
        /// </summary>
        /// <param name="refObject"></param>
        /// <param name="NameStage"></param>
        public void СменитьСтадию(ReferenceObject refObject, string NameStage)
        {
            TFlex.DOCs.Model.Stages.Stage Разработка = TFlex.DOCs.Model.Stages.Stage.Find(Connection, NameStage);
            Разработка.Change(new List<ReferenceObject> { refObject });
        }

        /// <summary>
        /// УдалитьСтадиюCОбъекта
        /// </summary>
        /// <param name="refObject"></param>
        public void УдалитьСтадиюCОбъекта(ReferenceObject refObject)
        {
            TFlex.DOCs.Model.Stages.Stage.Clear(Connection, new List<ReferenceObject> { refObject });
        }

        /// <summary>
        /// Получение описания справочника
        /// </summary>
        /// <param name="referenceGuid">Guid справочника, описание которого нужно получить</param>
        /// <returns></returns>
        public ReferenceInfo GetReferenceInfo(Guid referenceGuid)
        {
            ReferenceInfo referenceInfo = Connection.ReferenceCatalog.Find(referenceGuid);
            return referenceInfo;
        }

        /// <summary>
        /// Получить описание параметра 
        /// </summary>
        /// <param name="referenceInfo">Описание справочника</param>
        /// <param name="parameterGuid">Guid параметра</param>
        /// <returns>ParameterInfo</returns>
        public TFlex.DOCs.Model.Structure.ParameterInfo GetParameterInfo(ReferenceInfo referenceInfo, Guid parameterGuid)
        {
            TFlex.DOCs.Model.Structure.ParameterInfo parameterInfo = referenceInfo.Description.OneToOneParameters.Find(parameterGuid);
            return parameterInfo;
        }

        /// <summary>
        /// Получить класс справочника по Guid
        /// </summary>
        /// <param name="guidClass">Guid типа</param>
        /// <param name="reference">Справочник, в котором находится тип</param>
        /// <returns></returns>
        public ClassObject GetClassObject(Reference reference, Guid guidClass)
        {
            ClassObject classObject = reference.Classes.Find(guidClass);
            return classObject;
        }
    }


}//by Kachkov Ivan
