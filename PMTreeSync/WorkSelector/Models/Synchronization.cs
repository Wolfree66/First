using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Search;
//using TFlex.DOCs.Model.Types;

namespace TFLEXDocsClasses
{
    /// <summary>
    /// Для работы с синхронизацией в УП
    /// ver 1.3
    /// </summary>
    class Synchronization
    {
        /// <summary>
        /// Возвращает список синхронизированных работ из указанного пространства планирования
        /// </summary>
        /// <param name="work"></param>
        /// <param name="planningSpaceGuidString"></param>
        /// <param name="returnOnlyMasterWorks"> если true только укрупнения, если false только детализации</param>
        /// <returns></returns>
        public static List<ProjectManagementWork> GetSynchronizedWorksInProject(ProjectManagementWork work, ProjectManagementWork project, bool? returnOnlyMasterWorks = null)
        {
            string guidWorkForSearch = work.Guid.ToString();
            string guidProjectForSearch = project.Guid.ToString();
            Filter filter = new Filter(ProjectDependenciesReferenceInfo);
            // Условие: тип зависимости "синхронизация"
            ReferenceObjectTerm termSyncType = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
            // устанавливаем параметр
            termSyncType.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_DependencyType_Guid));
            // устанавливаем оператор сравнения
            termSyncType.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termSyncType.Value = 5;// DependencyType.Synchronization;

            // Условие: в названии содержится "является укрупнением"
            //ReferenceObjectTerm termMasterWork = new ReferenceObjectTerm(LogicalOperator.Or);

            ReferenceObjectTerm termMasterWork = new ReferenceObjectTerm(filter.Terms);
            // устанавливаем параметр
            termMasterWork.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_MasterWork_Guid));
            // устанавливаем оператор сравнения
            termMasterWork.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termMasterWork.Value = guidWorkForSearch;

            ReferenceObjectTerm termSlaveProjectWork = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
            // устанавливаем параметр
            termSlaveProjectWork.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_SlaveProject_Guid));
            // устанавливаем оператор сравнения
            termSlaveProjectWork.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termSlaveProjectWork.Value = guidProjectForSearch;

            // Условие: в названии содержится "является детализацией"
            ReferenceObjectTerm termSlaveWork = new ReferenceObjectTerm(filter.Terms, LogicalOperator.Or);
            //ReferenceObjectTerm term1 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.Or);
            // устанавливаем параметр
            termSlaveWork.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_SlaveWork_Guid));
            // устанавливаем оператор сравнения
            termSlaveWork.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termSlaveWork.Value = guidWorkForSearch;

            ReferenceObjectTerm termMasterProjectWork = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
            // устанавливаем параметр
            termMasterProjectWork.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_MasterProject_Guid));
            // устанавливаем оператор сравнения
            termMasterProjectWork.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termMasterProjectWork.Value = guidProjectForSearch;

            // Группируем условия в отдельную группу (другими словами добавляем скобки)
            TermGroup group1 = filter.Terms.GroupTerms(new Term[] { termMasterWork, termSlaveProjectWork });
            TermGroup group2 = filter.Terms.GroupTerms(new Term[] { termSlaveWork, termMasterProjectWork });

            //редактируем при необходимости
            if (returnOnlyMasterWorks != null)
            {
                if ((bool)returnOnlyMasterWorks) group1.Clear();
                else group2.Clear();
            }

            //Получаем список объектов, в качестве условия поиска – сформированный фильтр
            List<ReferenceObject> listObj = ProjectDependenciesReference.Find(filter);
            List<Guid> listSyncWorkGuids = new List<Guid>();
            //#if test
            //            System.Windows.Forms.MessageBox.Show(filter.ToString() + "\nlistObj.Count = " + listObj.Count.ToString());
            //#endif
            foreach (var item in listObj) // выбираем все отличные от исходной работы guids
            {
                Guid slaveGuid = item[SynchronizationParameterGuids.param_SlaveWork_Guid].GetGuid();
                Guid masterGuid = item[SynchronizationParameterGuids.param_MasterWork_Guid].GetGuid();
                if (slaveGuid != work.Guid) listSyncWorkGuids.Add(slaveGuid);
                if (masterGuid != work.Guid) listSyncWorkGuids.Add(masterGuid);
            }
            listSyncWorkGuids = listSyncWorkGuids.Distinct().ToList();
            List<ProjectManagementWork> result = new List<ProjectManagementWork>();

            foreach (var item in listSyncWorkGuids)
            {
                ProjectManagementWork tempWork = new ProjectManagementWork(item);
                result.Add(new ProjectManagementWork(item));

            }
            return result;


        }

        [Obsolete]
        /// <summary>
        /// Use GetSyncronizedWorks
        /// Возвращает список синхронизированных работ из указанного пространства планирования
        /// </summary>
        /// <param name="work"></param>
        /// <param name="planningSpaceGuidString"></param>
        /// <param name="returnOnlyMasterWorks"> если true только укрупнения, если false только детализации</param>
        /// <returns></returns>
        public static List<ProjectManagementWork> GetSynchronizedWorksFromSpace(ProjectManagementWork work, string planningSpaceGuidString, bool? returnOnlyMasterWorks = null)
        {
            string guidStringForSearch = work.Guid.ToString();
            Filter filter = GetFilter(returnOnlyMasterWorks, guidStringForSearch);

            //Получаем список объектов, в качестве условия поиска – сформированный фильтр
            List<ReferenceObject> listObj = ProjectDependenciesReference.Find(filter);
            List<Guid> listSyncWorkGuids = new List<Guid>();
            //#if test
            //            MessageBox.Show(filter.ToString() + "\nlistObj.Count = " + listObj.Count.ToString());
            //#endif
            foreach (var item in listObj) // выбираем все отличные от исходной работы guids
            {
                Guid slaveGuid = item[SynchronizationParameterGuids.param_SlaveWork_Guid].GetGuid();
                Guid masterGuid = item[SynchronizationParameterGuids.param_MasterWork_Guid].GetGuid();
                if (slaveGuid != work.Guid) listSyncWorkGuids.Add(slaveGuid);
                if (masterGuid != work.Guid) listSyncWorkGuids.Add(masterGuid);
            }
            listSyncWorkGuids = listSyncWorkGuids.Distinct().ToList();
            List<ProjectManagementWork> result = new List<ProjectManagementWork>();
            foreach (var item in listSyncWorkGuids)
            {
                ProjectManagementWork tempWork = new ProjectManagementWork(item);
                if (tempWork.PlanningSpace.ToString() == planningSpaceGuidString)
                {
                    result.Add(new ProjectManagementWork(item));
                }
            }
            return result;


        }

        /// <summary>
        /// Возвращает список синхронизированных работ из указанного пространства планирования
        /// </summary>
        /// <param name="work"></param>
        /// <param name="planningSpaceGuidString"></param>
        /// <param name="returnOnlyMasterWorks"> если true только укрупнения, если false только детализации</param>
        /// <returns></returns>
        public static List<ProjectManagementWork> GetSynchronizedWorks(ProjectManagementWork work, string planningSpaceGuidString, bool? returnOnlyMasterWorks)
        {
            string guidStringForSearch = work.Guid.ToString();
            Filter filter = GetFilter(returnOnlyMasterWorks, guidStringForSearch);

            //Получаем список объектов, в качестве условия поиска – сформированный фильтр
            List<ReferenceObject> listObj = ProjectDependenciesReference.Find(filter);
            List<Guid> listSyncWorkGuids = new List<Guid>();
            //#if test
            //            MessageBox.Show(filter.ToString() + "\nlistObj.Count = " + listObj.Count.ToString());
            //#endif
            foreach (var item in listObj) // выбираем все отличные от исходной работы guids
            {
                Guid slaveGuid = item[SynchronizationParameterGuids.param_SlaveWork_Guid].GetGuid();
                Guid masterGuid = item[SynchronizationParameterGuids.param_MasterWork_Guid].GetGuid();
                if (slaveGuid != work.Guid) listSyncWorkGuids.Add(slaveGuid);
                if (masterGuid != work.Guid) listSyncWorkGuids.Add(masterGuid);
            }
            listSyncWorkGuids = listSyncWorkGuids.Distinct().ToList();
            List<ProjectManagementWork> result = new List<ProjectManagementWork>();
            bool isNeedToFilterByPlanningSpace = !string.IsNullOrWhiteSpace(planningSpaceGuidString);
            foreach (var item in listSyncWorkGuids)
            {
                ProjectManagementWork tempWork = new ProjectManagementWork(item);
                if (isNeedToFilterByPlanningSpace && tempWork.PlanningSpace.ToString() != planningSpaceGuidString)
                {
                    continue;
                }
                result.Add(new ProjectManagementWork(item));

            }
            return result;

        }

        private static Filter GetFilter(bool? returnOnlyMasterWorks, string guidStringForSearch)
        {
            Filter filter = new Filter(ProjectDependenciesReferenceInfo);
            // Условие: тип зависимости "синхронизация"
            ReferenceObjectTerm termSyncType = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
            // устанавливаем параметр
            termSyncType.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_DependencyType_Guid));
            // устанавливаем оператор сравнения
            termSyncType.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termSyncType.Value = 5;

            // Условие: в названии содержится "является укрупнением"
            //ReferenceObjectTerm termMasterWork = new ReferenceObjectTerm(LogicalOperator.Or);

            ReferenceObjectTerm termMasterWork = new ReferenceObjectTerm(filter.Terms);
            // устанавливаем параметр
            termMasterWork.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_MasterWork_Guid));
            // устанавливаем оператор сравнения
            termMasterWork.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termMasterWork.Value = guidStringForSearch;
            // Условие: в названии содержится "является детализацией"
            ReferenceObjectTerm termSlaveWork = new ReferenceObjectTerm(filter.Terms, LogicalOperator.Or);
            //ReferenceObjectTerm term1 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.Or);
            // устанавливаем параметр
            termSlaveWork.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_SlaveWork_Guid));
            // устанавливаем оператор сравнения
            termSlaveWork.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termSlaveWork.Value = guidStringForSearch;

            // Группируем условия в отдельную группу (другими словами добавляем скобки)
            TermGroup group1 = filter.Terms.GroupTerms(new Term[] { termMasterWork, termSlaveWork });

            //редактируем при необходимости
            if (returnOnlyMasterWorks != null)
            {
                if (!(bool)returnOnlyMasterWorks) group1.Remove(termSlaveWork);
                else group1.Remove(termMasterWork);
            }

            return filter;
        }

        //создаём синхронизацию
        public static void SyncronizeWorks(ProjectManagementWork masterWork, ProjectManagementWork slaveWork)
        {
            ReferenceObject синхронизация = ProjectDependenciesReference.CreateReferenceObject(Class_ProjectDependency);
            синхронизация[SynchronizationParameterGuids.param_DependencyType_Guid].Value = 5;// DependencyType.Synchronization;
            синхронизация[SynchronizationParameterGuids.param_MasterWork_Guid].Value = masterWork.Guid;
            синхронизация[SynchronizationParameterGuids.param_SlaveWork_Guid].Value = slaveWork.Guid;
            синхронизация[SynchronizationParameterGuids.param_MasterProject_Guid].Value = masterWork.Project.Guid;
            синхронизация[SynchronizationParameterGuids.param_SlaveProject_Guid].Value = slaveWork.Project.Guid;
            синхронизация.EndChanges();
            //return синхронизация;
        }

        /// <summary>
        /// Удаляет синхронизацию между указанными работами
        /// </summary>
        /// <param name="masterWork"></param>
        /// <param name="slaveWork"></param>
        public static void DeleteSynchronisationBetween(ProjectManagementWork masterWork, ProjectManagementWork slaveWork)
        {

            Filter filter = new Filter(ProjectDependenciesReferenceInfo);
            // Условие: тип зависимости "синхронизация"
            ReferenceObjectTerm termSyncType = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
            // устанавливаем параметр
            termSyncType.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_DependencyType_Guid));
            // устанавливаем оператор сравнения
            termSyncType.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termSyncType.Value = 5;

            // Условие: в названии содержится "является укрупнением"
            string masterWorkGuidStringForSearch = masterWork.Guid.ToString();

            ReferenceObjectTerm termMasterWork = new ReferenceObjectTerm(filter.Terms);
            // устанавливаем параметр
            termMasterWork.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_MasterWork_Guid));
            // устанавливаем оператор сравнения
            termMasterWork.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termMasterWork.Value = masterWorkGuidStringForSearch;

            // Условие: в названии содержится "является детализацией"
            string slaveWorkGuidStringForSearch = slaveWork.Guid.ToString();
            ReferenceObjectTerm termSlaveWork = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
            //ReferenceObjectTerm term1 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.Or);
            // устанавливаем параметр
            termSlaveWork.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_SlaveWork_Guid));
            // устанавливаем оператор сравнения
            termSlaveWork.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termSlaveWork.Value = slaveWorkGuidStringForSearch;


            //Получаем список объектов, в качестве условия поиска – сформированный фильтр
            List<ReferenceObject> listObj = ProjectDependenciesReference.Find(filter);
            //#if test
            //            MessageBox.Show(filter.ToString() + "\nlistObj.Count = " + listObj.Count.ToString());
            //#endif
            foreach (var item in listObj) // удаляем всё что нашли
            {
                item.Delete();
            }
        }


        private static ClassObject _Class_ProjectDependency;
        /// <summary>
        /// класс Запрос коррекции работ
        /// </summary>
        public static ClassObject Class_ProjectDependency
        {
            get
            {
                if (_Class_ProjectDependency == null)
                    _Class_ProjectDependency = ProjectDependenciesReference.Classes.Find(SynchronizationParameterGuids.class_Dependency_GUID);

                return _Class_ProjectDependency;
            }
        }


        static Reference _ProjectDependenciesReference;
        /// <summary>
        /// Справочник "Зависимости проектов"
        /// </summary>
        public static Reference ProjectDependenciesReference
        {
            get
            {
                if (_ProjectDependenciesReference == null)
                    _ProjectDependenciesReference = GetReference(ref _ProjectDependenciesReference, ProjectDependenciesReferenceInfo);

                return _ProjectDependenciesReference;
            }
        }

        static ReferenceInfo _ProjectDependenciesReferenceInfo;
        /// <summary>
        /// Справочник "Зависимости проектов"
        /// </summary>
        private static ReferenceInfo ProjectDependenciesReferenceInfo
        {
            get { return GetReferenceInfo(ref _ProjectDependenciesReferenceInfo, SynchronizationParameterGuids.ref_ProjDependencies_Guid); }
        }

        private static Reference GetReference(ref Reference reference, ReferenceInfo referenceInfo)
        {
            if (reference == null)
                reference = referenceInfo.CreateReference();


            return reference;
        }
        private static ReferenceInfo GetReferenceInfo(ref ReferenceInfo referenceInfo, Guid referenceGuid)
        {
            if (referenceInfo == null)
                referenceInfo = References.Connection.ReferenceCatalog.Find(referenceGuid);

            return referenceInfo;
        }
        struct SynchronizationParameterGuids
        {
            //Справочник - "Зависимости проектов"
            public static readonly Guid ref_ProjDependencies_Guid = new Guid("e13cee45-39fa-43ff-ba2a-957294d975bf");    //Guid справочника - "Зависимости проектов"
            public static readonly Guid class_Dependency_GUID = new Guid("0ed0174e-5392-460f-8b53-5a2e52c26f9b");       //Guid типа "Зависимости проектов" справочника - "Зависимости проектов"
            public static readonly Guid param_MasterProject_Guid = new Guid("087648ca-6269-4e7f-8e5a-9fabbab5fafd");    //Guid параметра "Проект 1"
            public static readonly Guid param_MasterWork_Guid = new Guid("17feb793-7656-475a-b7b0-61672ba813fa");    //Guid параметра "Работа 1"
            public static readonly Guid param_SlaveProject_Guid = new Guid("a5590629-9469-4c39-87e6-169b60678abb");    //Guid параметра "Проект 2"
            public static readonly Guid param_SlaveWork_Guid = new Guid("5baf8272-f639-4fc8-af65-6f0af32647fd");    //Guid параметра "Работа 2"
            public static readonly Guid param_DependencyType_Guid = new Guid("45bc3bf0-88cf-4b46-b951-b1e07a5d7fcb");    //Guid параметра "Тип зависимости"
        }
    }
}
