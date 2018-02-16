using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using WpfApp_DialogueAddSignatories.Model;
using TFlexDOCsPM = TFlex.DOCs.References.ProjectManagement;

namespace WpfApp_DialogueAddSignatories.Model
{
    /// <summary>
    /// Информационный класс Элемент проекта - Управления проектами
    /// ver 1.11
    /// </summary>
    public class ProjectManagementWork : ProjectTreeItem
    {


        public ProjectManagementWork(Guid guid)
        {
            Guid = guid;
            ReferenceObject work = FindWork(Guid);
            ReferenceObject = work;

        }

        public ProjectManagementWork(string stringGuid) : this(new Guid(stringGuid))
        { }

        public ProjectManagementWork(ReferenceObject work)
        {
            ReferenceObject = work;
            this.Guid = work.SystemFields.Guid;
        }



        // ReferenceObject _ReferenceObject;
        //public ReferenceObject ReferenceObject
        //{
        //    get
        //    {
        //        return _ReferenceObject;
        //    }
        //    set
        //    {
        //        if (_ReferenceObject != value)
        //        {
        //            _ReferenceObject = value;
        //            if (_ReferenceObject != null)
        //            {
        //                LastEditDate = _ReferenceObject.SystemFields.EditDate;
        //            }

        //        }
        //    }
        //}

        public Guid Guid { get; private set; }

        string _Number;
        /// <summary>
        /// номер работы
        /// </summary>
        public string Number
        {
            get
            {
                if (this.IsProject || this.ParentWork == null)
                { _Number = ""; }
                else if (this.ParentWork != null)
                {
                    if (string.IsNullOrWhiteSpace(this.ParentWork.Number))
                        _Number = this.ReferenceObject.SystemFields.Order.ToString();
                    else
                        _Number = this.ParentWork.Number + "." + this.ReferenceObject.SystemFields.Order.ToString();
                }
                return _Number;
            }
        }


        /// <summary>
        /// Находит синхронизированную работу из пространства Рабочие планы двигаясь вверх по дереву от указанной работы
        /// </summary>
        /// <param name="work"></param>
        /// <returns>работа из пространства Рабочие планы, если не найдена то null</returns>
        private ProjectManagementWork GetTheWorkFromWorkingPlanningSpace(ProjectManagementWork work)
        {
            ProjectManagementWork result = null;
            if (work == null) return result;
            //MacroContext mc = new MacroContext(ServerGateway.Connection);
            // MessageBox.Show(work.ToString());
            if (work.PlanningSpace.ToString() == PlanningSpaceGuidString.WorkingPlans)
            {
                //MessageBox.Show("work from PlanningSpaceGuidString" + work.ToString());
                return work;
            }
            else
            {
                foreach (var item in Synchronization.GetSynchronizedWorksFromSpace(work, PlanningSpaceGuidString.WorkingPlans))
                //foreach (var item in Synchronization.GetSynchronizedWorksFromSpace(work, PlanningSpaceGuidString.WorkingPlans))
                {
                    while (result == null)
                    {
                        result = GetTheWorkFromWorkingPlanningSpace(item);
                    }
                }
            }
            if (result == null && work.ParentWork != null) result = GetTheWorkFromWorkingPlanningSpace(work.ParentWork);
            return result;
        }

        /// <summary>
        /// Находит кубик из 1С, если не получилось то null
        /// </summary>
        /// <param name="planningSpaceWork"></param>
        /// <returns>кубик из 1С, если не получилось то null</returns>
        private ProjectManagementWork GetTheProjectFrom1C(ProjectManagementWork planningSpaceWork)
        {

            //MessageBox.Show("GetTheProjectFrom1C" + planningSpaceWork.PlanningSpace.ToString() + "\nIsProjectFrom1C = " + planningSpaceWork.IsProjectFrom1C.ToString()+"\n" + planningSpaceWork.Name);
            if (planningSpaceWork.PlanningSpace.ToString() != PlanningSpaceGuidString.WorkingPlans)
            {

                //MessageBox.Show(planningSpaceWork.PlanningSpace.ToString() + " != " + PlanningSpaceGuidString.WorkingPlans);
                planningSpaceWork = GetTheWorkFromWorkingPlanningSpace(planningSpaceWork);
            }
            else
            {
                if (planningSpaceWork.IsProjectFrom1C) return planningSpaceWork;
            }
            //if (planningSpaceWork.Class.Guid == PM_class_Project_Guid) return planningSpaceWork;
            if (planningSpaceWork == null)
            {
                //MessageBox.Show("planningSpaceWork == null");
                return null;
            }
            //MessageBox.Show("GetTheParentProjectFrom1C" + planningSpaceWork.PlanningSpace.ToString() + planningSpaceWork.Name);
            return GetTheParentProjectFrom1C(planningSpaceWork);
        }

        private ProjectManagementWork GetTheParentProjectFrom1C(ProjectManagementWork work)
        {
            if (work.IsProjectFrom1C) return work;
            //if (planningSpaceWork.Class.Guid == PM_class_Project_Guid) return planningSpaceWork;
            if (work.ParentWork == null) return null;
            return GetTheParentProjectFrom1C(work.ParentWork);
        }

        private ReferenceObject FindWork(Guid guid)
        {
            //Guid PMReferenceGuid = new Guid("86ef25d4-f431-4607-be15-f52b551022ff");
            //Reference PMReference = ServerGateway.Connection.ReferenceCatalog.Find(PMReferenceGuid).CreateReference();
            Reference PMReference = References.ProjectManagementReference;
            ReferenceObject work = PMReference.Find(guid);
            return work;
        }

        string _Status;
        public string Status
        {
            get
            {
                if (_Status == null)
                {
                    string result = "";
                    TFlexDOCsPM.ProjectElement current;
                    if (ReferenceObject is TFlex.DOCs.References.ProjectManagement.ProjectElement) current = ReferenceObject as TFlexDOCsPM.ProjectElement;
                    else return "";

                    double Percent = current[PM_param_Percent_GUID].GetDouble();

                    DateTime PlanStartDate = current[PM_param_PlanStartDate_GUID].GetDateTime();
                    DateTime PlanEndDate = current[PM_param_PlanEndDate_GUID].GetDateTime();
                    DateTime FactStartDate = current[PM_param_FactStartDate_GUID].GetDateTime();
                    bool FactStartDateIsEmpty = current[PM_param_FactStartDate_GUID].IsEmpty;
                    DateTime FactEndDate = current[PM_param_FactEndDate_GUID].GetDateTime();
                    bool FactEndDateIsEmpty = current[PM_param_FactEndDate_GUID].IsEmpty;

                    result = "Данные не корректны.";
                    /*DateTime времяноль = new DateTime();*/
                    if (Percent == 1) return "Выполнено";
                    if (PlanStartDate.Date.CompareTo(DateTime.Now.Date) > 0 && FactStartDateIsEmpty && Percent == 0) //плановое начало позже сегодняшнего числа
                    {
                        result = "Срок не наступил";
                    }
                    else //плановое начало (уже прошло) раньше сегодняшнего числа
                    {
                        if (FactStartDateIsEmpty && Percent == 0) result = "Не начиналось";
                        if (PlanEndDate.Date.CompareTo(DateTime.Now.Date) >= 0)//плановое окончание позже или сегодняшнего числа (ещё не наступило)
                        {
                            if ((!FactStartDateIsEmpty && FactEndDateIsEmpty) ||
                                (Percent > 0 && Percent < 1)) // 0 < % < 1
                                result = "В работе";

                        }
                        else //плановое окончание раньше сегодняшнего числа (уже прошло)
                        {
                            if (/*!current[factStart].IsEmpty && */FactEndDateIsEmpty && //есть факт. старт и нет факт. окончания
                                Percent != 1) //процент выполнения не 100%
                                result = "Не выполнена";

                        }
                    }

                    return result;
                }
                return _Status;
            }
        }

        string _Name;
        public string Name
        {
            get
            {
                if (_Name == null)
                {
                    if (ReferenceObject != null)
                        _Name = ReferenceObject[PM_param_Name_GUID].GetString();
                    else _Name = "";
                }
                return _Name;
            }
        }

        DateTime? _StartDate;
        public DateTime StartDate
        {
            get
            {
                if (_StartDate == null)
                {
                    if (ReferenceObject != null)
                        _StartDate = ReferenceObject[PM_param_PlanStartDate_GUID].GetDateTime();
                    else _StartDate = DateTime.MinValue;
                }
                return (DateTime)_StartDate;
            }
            private set
            {
                if (_StartDate != value)
                    _StartDate = value;
            }
        }
        DateTime? _EndDate;
        public DateTime EndDate
        {
            get
            {
                if (_EndDate == null)
                {
                    if (ReferenceObject != null)
                        _EndDate = ReferenceObject[PM_param_PlanEndDate_GUID].GetDateTime();
                    else _EndDate = DateTime.MinValue;
                }
                return (DateTime)_EndDate;
            }
            private set
            {
                if (_EndDate != value)
                    _EndDate = value;
            }
        }

        //string _UsedResourcesNames;
        public string UsedResourcesNames
        {
            get
            {
                string result = "";
                if (PlannedNonConsumableResources != null)
                {
                    int i = 0;
                    foreach (var item in PlannedNonConsumableResources)
                    {
                        if (i == 0) result += item.Resource.ToString();
                        else result += Environment.NewLine + item.Resource.ToString();
                        i++;
                    }
                }
                return result;
            }
        }

        List<UsedResource> _PlannedNonConsumableResources;
        public IEnumerable<UsedResource> PlannedNonConsumableResources
        {
            get
            {
                if (_PlannedNonConsumableResources == null)
                {
                    List<ReferenceObject> resources = this.ReferenceObject.GetObjects(ProjectManagementWork.PM_link_UsedResources_GUID).Where(res => (res.Class == References.Class_NonConsumableResources)).ToList();
                    _PlannedNonConsumableResources = new List<UsedResource>();
                    foreach (var item in resources)
                    {
                        UsedResource usedResource = new UsedResource(item);
                        if (usedResource.IsPlanned) _PlannedNonConsumableResources.Add(usedResource);
                    }
                }
                return _PlannedNonConsumableResources;
            }
        }

        double? _LabourCost;
        public double LabourCost
        {
            get
            {
                if (_LabourCost == null)
                {
                    if (ReferenceObject != null)
                    { _LabourCost = this.ReferenceObject[PM_param_SumPlanResourceCount_GUID].GetDouble(); }
                    else _LabourCost = 0;
                }
                return (double)_LabourCost;
            }
        }

        string _ProjectCipherFrom1C;
        public string ProjectCipherFrom1C
        {
            get
            {
                if (_ProjectCipherFrom1C == null && ProjectFrom1C != null)
                {
                    //MessageBox.Show("ProjectFrom1C -" + ProjectFrom1C.ToString());
                    _ProjectCipherFrom1C = ProjectFrom1C.ReferenceObject[PM_param_Cipher_GUID].GetString();
                }
                if (ProjectFrom1C == null)
                {
                    //MessageBox.Show("ProjectFrom1C == null");
                    _ProjectCipherFrom1C = "Не указан";
                }
                return _ProjectCipherFrom1C;
            }
        }

        string _ProjectNumberFrom1C;
        public string ProjectNumberFrom1C
        {
            get
            {
                if (_ProjectNumberFrom1C == null && ProjectFrom1C != null)
                {
                    //MessageBox.Show("ProjectFrom1C -" + ProjectFrom1C.ToString());
                    _ProjectNumberFrom1C = ProjectFrom1C.ReferenceObject[PM_param_Name_GUID].GetString();
                }
                if (ProjectFrom1C == null)
                {
                    //MessageBox.Show("ProjectFrom1C == null");
                    _ProjectNumberFrom1C = "Не указан";
                }
                return _ProjectNumberFrom1C;
            }
        }

        string _ProjectResponsibleNameFrom1C;
        public string ProjectResponsibleNameFrom1C
        {
            get
            {
                if (_ProjectResponsibleNameFrom1C == null && ProjectFrom1C != null)
                {
                    ReferenceObject responsible = ProjectFrom1C.ReferenceObject.GetObject(PM_link_Responsible_GUID);
                    if (responsible != null)
                        _ProjectResponsibleNameFrom1C = responsible.ToString();
                }
                if (_ProjectResponsibleNameFrom1C == null)
                    _ProjectResponsibleNameFrom1C = "Не указан";
                return _ProjectResponsibleNameFrom1C;
            }

        }

        Guid? _PlanningSpace;
        public Guid PlanningSpace
        {
            get
            {
                if (_PlanningSpace == null)
                {
                    if (Project != null)
                        _PlanningSpace = Project.ReferenceObject[PM_param_PlanningSpace_GUID].GetGuid();
                    else _PlanningSpace = Guid.Empty;
                }
                return (Guid)_PlanningSpace;
            }
        }

        List<ProjectManagementWork> _Children;
        public new List<ProjectManagementWork> Children
        {
            get
            {
                if (_Children == null && this.ReferenceObject.Children != null)
                {
                    _Children = new List<ProjectManagementWork>();
                    foreach (var child in this.ReferenceObject.Children)
                    {
                        _Children.Add(Factory.Create_ProjectManagementWork(child));
                    }
                }
                return _Children;
            }
        }

        ProjectManagementWork _ParentWork;
        public ProjectManagementWork ParentWork
        {
            get
            {
                if (_ParentWork == null && this.ReferenceObject.Parent != null)
                    if (this.ReferenceObject.Parent.Class.IsInherit(PM_class_ProjectElement_Guid))
                        _ParentWork = new ProjectManagementWork/* Factory.Create_ProjectManagementWork*/(this.ReferenceObject.Parent);
                return _ParentWork;
            }
        }

        bool? _IsProjectFrom1C;
        public bool IsProjectFrom1C
        {
            get
            {

                if (_IsProjectFrom1C == null)
                {
                    _IsProjectFrom1C = this.ReferenceObject.Class.IsInherit(PM_class_ProjectFrom1C_Guid);
                }
                return (bool)_IsProjectFrom1C;
            }
        }

        /// <summary>
        /// true если ReferenceObject не существует, в противном случае false
        /// </summary>
        public bool IsEmpty
        {
            get
            {
                return ReferenceObject == null;
            }
        }

        bool? _IsProject;
        public bool IsProject
        {
            get
            {

                if (_IsProject == null)
                {
                    _IsProject = this.ReferenceObject.Class.IsInherit(PM_class_Project_Guid);
                }
                return (bool)_IsProject;
            }
        }

        ProjectManagementWork GetTheProject(ProjectManagementWork work)
        {
            //MessageBox.Show("GetTheProject\n" + work.ToString() + "\n" + work.Class.Guid.ToString() + "\n" + PM_class_Project_Guid.ToString());
            if (work.IsProject) return work;
            if (work.ParentWork == null)
            {
                //MessageBox.Show("work.Parent == null");
                return work;
            }
            //MessageBox.Show("work.Parent == " + work.Parent.ToString());
            return GetTheProject(work.ParentWork);
        }
        public ServerConnection Connection { get { return ServerGateway.Connection; } }

        TFlex.DOCs.Model.References.Users.User _Answerable;
        public TFlex.DOCs.Model.References.Users.User Answerable
        {
            get
            {
                if (_Answerable == null)
                {
                    _Answerable = this.ReferenceObject.GetObject(PM_link_Responsible_GUID) as TFlex.DOCs.Model.References.Users.User;
                }
                return _Answerable;
            }
        }

        ProjectManagementWork _ProjectFrom1C;

        public ProjectManagementWork ProjectFrom1C
        {
            get
            {
                if (_ProjectFrom1C == null)
                { _ProjectFrom1C = GetTheProjectFrom1C(this); }
                return _ProjectFrom1C;
            }
        }

        ProjectManagementWork _Project;

        internal ProjectManagementWork Project
        {
            get
            {
                if (_Project == null)
                { _Project = GetTheProject(this); }
                return _Project;
            }
        }

        string _ProjectName;
        public string ProjectName
        {
            get
            {
                if (_ProjectName == null)
                    _ProjectName = this.Project.ToString();
                if (_ProjectName == null)
                    _ProjectName = "Нет проекта";
                return _ProjectName;
            }
        }

        /*List<ReferenceObject> _PlannedUsedResources;
        public IEnumerable<ReferenceObject> PlannedUsedResources
        {
            get
            {
                if (_PlannedUsedResources == null)
                {
                    var usedRes = ReferenceObject.GetObjects(PM_link_UsedResources_GUID).Where(res => (res as ResourcesUsedReferenceObject).ResourcesLink is NonExpendableResourcesReferenceObject);
                    if (usedRes != null)
                    {
                        _PlannedUsedResources = usedRes.Distinct().ToList();
                    }
                    else _PlannedUsedResources = new List<ReferenceObject>();
                    //MessageBox.Show(usedRes.Count().ToString());
                }
                //MessageBox.Show(_PlannedUsedResources.Count().ToString());
                return _PlannedUsedResources;
            }
        }
        */
        double? _Progress;
        public int Progress
        {
            get
            {
                if (_Progress == null)
                {
                    _Progress = ReferenceObject[PM_param_Percent_GUID].GetDouble() * 100;
                }
                return (int)_Progress;
            }
        }

        public DateTime LastEditDate { get; internal set; }

        public override string ToString()
        {
            return this.Name;
        }

        static public ObservableCollection<ProjectManagementWork> AllDetailingProjects(List<ReferenceObject> ЗависимостиДетализации)
        {
            ObservableCollection<ProjectManagementWork> detailingProjects = new ObservableCollection<ProjectManagementWork>();
            Guid guidЗависимости = new Guid();
            var PMref = WpfApp_DialogueAddSignatories.Model.References.ProjectManagementReference;
            foreach (var зависимость in ЗависимостиДетализации)
            {

                guidЗависимости = зависимость[Synchronization.SynchronizationParameterGuids.param_SlaveProject_Guid].GetGuid();

                //var pe = PMref.Find(guidЗависимости);

                var PMW = new ProjectManagementWork(guidЗависимости);

                detailingProjects.Add(PMW);
            }
            return detailingProjects;
        }



        /// <summary>
        /// Изменяет даты начала и окончания на указанные
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        internal void ChangeDates(DateTime startDate, DateTime endDate)
        {
            BeginChanges();
            //if (this.EndDate <= startDate)
            //{
            //    this.ReferenceObject[PM_param_PlanEndDate_GUID].Value = endDate;
            //    this.ReferenceObject[PM_param_PlanStartDate_GUID].Value = startDate;
            //}
            //else// if (this.StartDate >= endDate)
            //{
            this.ReferenceObject[PM_param_PlanStartDate_GUID].Value = startDate;
            this.ReferenceObject.ApplyChanges();
            this.ReferenceObject[PM_param_PlanEndDate_GUID].Value = endDate;
            //}

            this.ReferenceObject.EndChanges(); //если нет изменений то возвращает false

            this.EndDate = endDate;
            this.StartDate = startDate;
        }

        /// <summary>
        /// Изменяет наименование работы
        /// </summary>
        /// <param name="name"></param>
        internal void ChangeWorkName(string name)
        {
            BeginChanges();
            this.ReferenceObject[PM_param_Name_GUID].Value = name;
            this.ReferenceObject.EndChanges(); //если нет изменений то возвращает false
            this._Name = name;
        }


        /// <summary>
        /// получить пользователей которые имеют доступ на элемент проекта "без ограничений"
        /// </summary>
        /// <param name="projectElement"></param>
        /// <param name="accessCode"></param>
        /// <returns></returns>
        public List<TFlex.DOCs.Model.References.Users.User> GetUsersHaveAccess(AccessCode accessCode)
        {
            /// <summary>
            ///  Guid типа "Доступы"
            ///  тип "Список объектов"
            /// </summary>
            Guid TypeAccesses_AccessesList_Guid = new Guid("68989495-719e-4bf3-ba7c-d244194890d5");

            Guid Accesses_TypeParameters_Access_Guid = new Guid("aff97f60-58d6-4c69-807b-dd7b6abfd770"); //Доступ Целое число 

            /// <summary>
            /// Guid параметра - "Уникальный идентификатор пользователя"
            /// Уникальный идентификатор
            /// </summary>
            Guid Accesses_TypeParameters_UserGuid_Guid = new Guid("f65f1d07-275e-44cd-9f9e-5492971e7ebf");
            List<TFlex.DOCs.Model.References.Users.User> users = new List<TFlex.DOCs.Model.References.Users.User>();
            foreach (ReferenceObject access in this.ReferenceObject.GetObjects(TypeAccesses_AccessesList_Guid))//перебор объектов
            {

                if (access[Accesses_TypeParameters_Access_Guid].GetInt32() == (int)accessCode
                    && access[Accesses_TypeParameters_UserGuid_Guid].GetGuid() != Guid.Empty)//Параметры
                {
                    TFlex.DOCs.Model.References.Users.User user = References.UsersReference.Find(access[Accesses_TypeParameters_UserGuid_Guid].GetGuid()) as TFlex.DOCs.Model.References.Users.User;

                    if (!users.Contains(user))
                        users.Add(user);
                }
            }

            return users;
        }


        /// <summary>
        /// переводит объект в состояние редактирования
        /// </summary>
        private void BeginChanges()
        {
            this.ReferenceObject.Reload();//без этой строки вылетает ошибка при взятии на редактирование незаблокированного объекта

            // if (this.ReferenceObject == null) this.ReferenceObject = References.WorkChangeRequestReference.CreateReferenceObject(References.Class_WorkChangeRequest);
            if (this.ReferenceObject == null) return;
            else if (this.ReferenceObject.LockState == ReferenceObjectLockState.LockedByCurrentUser)
            {
                this.ReferenceObject.BeginChanges();
            }
            else if (this.ReferenceObject.LockState == ReferenceObjectLockState.LockedByOtherUser)
            {
                this.ReferenceObject.Unlock();
                this.ReferenceObject.BeginChanges();
            }
            else if (this.ReferenceObject.LockState == ReferenceObjectLockState.None)
            {
                this.ReferenceObject.BeginChanges();
            }
            if (!this.ReferenceObject.CanEdit) throw new Exception("Редактирование работы не возможно.");
        }

        internal void Delete()
        {
            this.ReferenceObject.Delete();
            this.ReferenceObject = null;
        }

        internal void AddResource(UsedResource newUsedResource)
        {
            this.BeginChanges();
            newUsedResource.ReferenceObject.Reload();
            this.ReferenceObject.AddLinkedObject(ProjectManagementWork.PM_link_UsedResources_GUID, newUsedResource.ReferenceObject);
            this.ReferenceObject.EndChanges();
            _PlannedNonConsumableResources = null;
        }

        public enum AccessCode
        {
            OnlyViewing,
            NoLimit,
            Parameters,
            Resources = 4,
            ParametersAndResources = 6,
            Detalisation = 8,
            ParametersAndDetalisation = 10,
            ResourcesAndDetalisation = 12,
            ChildWorks = 16,
            ResourcesAndChildWorks = 20,
            DetalisationAndChildWorks = 24
        }

        internal static List<Guid> listParametersToSkip = new List<Guid>() {
        PM_param_PlanningSpace_GUID,//обязательные к пропуску параметры
        PM_param_WorkIntervalLengthDays_Guid, PM_param_WorkIntervalLengthHours_Guid, PM_param_FactStartDate_GUID, PM_param_FactEndDate_GUID,
        PM_param_ProjectCodeFrom1C_Guid};

        internal static List<Guid> listLinksToSkip = new List<Guid>() { PM_list_Access_GUID };

        public static Guid PM_class_ProjectElement_Guid = new Guid("7c968a5b-d1a4-468f-a8cc-d9dbf7ecc990"); //тип - "Элемент проекта"
        public static Guid PM_class_Work_Guid = new Guid("c0bef497-cf64-44a7-9839-a704dc3facb2"); //тип - "Работа"

        static Guid PM_class_ProjectFrom1C_Guid = new Guid("859f7412-c95c-4636-b1f0-f792392ccf83"); //тип - "Затраты проекта 1С new"
        static Guid PM_class_Project_Guid = new Guid("9906a1e9-dc3b-4b27-a083-b3dba9ee8ed0"); // тип - "Проект"
        static Guid PM_list_Access_GUID = new Guid("68989495-719e-4bf3-ba7c-d244194890d5");       //Guid списка "Доступы", справочника - "Управление проектами"
        static Guid PM_param_Name_GUID = new Guid("b7e3b3fe-65c0-4f1f-82e5-3ee95fe360dd");//Guid параметра "Наименование" справочника - "Управление проектами"
        static Guid PM_param_Cipher_GUID = new Guid("b4f4b4d3-5ffc-4869-a880-7561e0c2a574");//Guid параметра "Обозначение" справочника - "Управление проектами"
        static Guid PM_param_PlanningSpace_GUID = new Guid("fd123c69-7945-487a-a854-884ef68f1036");//Guid параметра "Пространство планирования" справочника - "Управление проектами"
        static Guid PM_param_PlanStartDate_GUID = new Guid("4fd8d58a-04ba-4e80-a240-12691e460b83");//Guid параметра "Начало"
        static Guid PM_param_PlanEndDate_GUID = new Guid("da997729-07e2-40aa-a1a0-c08bc4dde481");//Guid параметра "Окончание"
        static Guid PM_param_FactStartDate_GUID = new Guid("2f457df1-246d-4d23-a332-53f387940ba9");      //Guid параметра "Фактическое начало", справочника - "Управление проектами"
        static Guid PM_param_FactEndDate_GUID = new Guid("a93f9644-3e24-4a59-9eb9-d200c83044ee");//Guid параметра "Фактическое окончание"
        static Guid PM_param_UsedResourceIsFact_GUID = new Guid("4227e515-66a4-43ae-9418-346854748986");//Guid параметра "Фактическое значение"
        static Guid PM_param_Percent_GUID = new Guid("f1c3091f-2fff-48d7-a962-cf8d6b08bdd2");//Guid параметра "Процент выполнения"
        static Guid PM_param_ResultOfWork_GUID = new Guid("a24fe48c-cdb4-428f-9330-72fef743b713");//Guid параметра "Результат работы"
        static Guid PM_param_SumPlanResourceCount_GUID = new Guid("cde1942c-ffc6-4654-9a2a-8f0b3a0a8e14");//Guid параметра "Суммарные плановые трудозатраты по ресурсам"
        static Guid PM_param_ProjectCodeFrom1C_Guid = new Guid("b37cfad2-bd50-47f3-8292-8f7c10d8b9a2"); // Guid параметра "Код проекта из 1С", справочника - "Управление проектами"
        static Guid PM_param_WorkIntervalLengthDays_Guid = new Guid("488646b9-a912-468f-a791-e9e68774215f");      //Guid параметра "Длительность(Дни)", справочника - "Управление проектами"
        static Guid PM_param_WorkIntervalLengthHours_Guid = new Guid("9f733670-5f3b-4554-b023-e4d450ab0c00");      //Guid параметра "Длительность", справочника - "Управление проектами"

        public static readonly Guid PM_link_UsedResources_GUID = new Guid("1a22ee46-5438-4caa-8b75-8a8a37b74b7e");//Guid списка 1:n "Ресурс"
        static Guid PM_link_Responsible_GUID = new Guid("063df6fa-3889-4300-8c7a-3ce8408a931a");//Guid связи N:1 "Ответственный"

    }
}
