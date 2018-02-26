using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.References.ProjectManagement;
using WpfApp_DialogueAddSignatories.Infrastructure;
using System.Windows.Forms;
using System.Collections.ObjectModel;
using System.Windows.Input;
using WpfApp_DialogueAddSignatories.Model;

namespace WpfApp_DialogueAddSignatories.ViewModel
{
    public class MainWindowViewModel : ViewModelBase
    {


        private ReferenceObject startRefObject;

        private static ProjectManagementWork startObject;

        public ProjectManagementWork StartObject
        {
            get
            {
                if (startObject == null)
                    startObject = Factory.Create_ProjectManagementWork(startRefObject);
                return startObject;
            }
            private set
            {
                startObject = value;
                OnPropertyChanged("StartObject");
            }
        }



        public MainWindowViewModel(ReferenceObject startRefObject)
        {
            this.startRefObject = startRefObject;

            if (startObject == null)
                startObject = Factory.Create_ProjectManagementWork(startRefObject);

            if (IsListNullOrEmpty(DetailingProjects))
                ShowError("Синхронизация состава работа", "Ошибка, выбранная детализация не найдена!");

            return;
        }

        #region  

        private static bool _IsSyncRes = false;

        public bool IsSyncRes
        {
            get { return _IsSyncRes; }
            set { OnPropertyChanged("IsSyncRes"); _IsSyncRes = value; }
        }
        private static bool _IsSyncOnlyPlanRes = false;

        public bool IsSyncOnlyPlanRes
        {
            get { return _IsSyncOnlyPlanRes; }
            set { OnPropertyChanged("IsSyncOnlyPlanRes"); _IsSyncOnlyPlanRes = value; }
        }

        private object selectedDetailingProject;
        public object SelectedDetailingProject
        {
            get { return selectedDetailingProject; }
            set { selectedDetailingProject = value; OnPropertyChanged("SelectedDetailingProject"); }
        }

        #endregion


        ObservableCollection<ProjectManagementWork> _detailingProjects;
        /// <summary>
        /// Список проектов (детализаций) выбронного элемента проекта
        /// </summary>
        public ObservableCollection<ProjectManagementWork> DetailingProjects
        {
            get
            {
                if (IsListNullOrEmpty(_detailingProjects))
                {
                    _detailingProjects = LoadDetailingProjects();

                }
                return _detailingProjects;
            }
        }

        ObservableCollection<ProjectManagementWork> LoadDetailingProjects()
        {
            return ProjectManagementWork.AllDetailingProjects(ЗависимостиДетализации);
        }

        private static List<ReferenceObject> _зависимостиДетализации;

        public List<ReferenceObject> ЗависимостиДетализации
        {
            get
            {
                if (IsListNullOrEmpty(_зависимостиДетализации))
                {
                    _зависимостиДетализации = new List<ReferenceObject>();
                    _зависимостиДетализации = Synchronization.GetDependenciesObjects(false, StartObject);
                }
                return _зависимостиДетализации;
            }
        }

        void ShowError(string caption, string message)
        {
            System.Windows.Forms.MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1, System.Windows.Forms.MessageBoxOptions.ServiceNotification);
        }

        public static bool IsListNullOrEmpty<T>(IEnumerable<T> list)
        {
            if (list == null || list.Count() == 0)
                return true;

            return false;
        }


        public object Show { get; internal set; }

        /// <summary>
        /// Дерево 
        /// </summary>

        ObservableCollection<ProjectTreeItem> _tree = new ObservableCollection<ProjectTreeItem>();
        /// <summary>
        /// Список проектов (детализаций) выбронного элемента проекта
        /// </summary>
        public ObservableCollection<ProjectTreeItem> Tree
        {
            get
            {

                return _tree;
            }
        }

        //ProjectTreeItem tree;
        //public ProjectTreeItem Tree
        //{
        //    get
        //    {
        //        if (tree == null)
        //            tree = Factory.CreateProjectTreeItem(startRefObject) as ProjectTreeItem;
        //        return tree;
        //    }
        //}

        #region Commands



        RelayCommand _syncResCommand;
        /// <summary>
        /// Синхронизировать ресурсы
        /// </summary>
        public ICommand SyncResCommand
        {
            get
            {
                if (_syncResCommand == null)
                    _syncResCommand = new RelayCommand(ExecuteSyncResCommand, CanExecuteSyncResCommand);
                return _syncResCommand;
            }
        }

        /// <summary>
        /// Смена флага на противоположное значение
        /// </summary>
        /// <param name="parameter"></param>
        public void ExecuteSyncResCommand(object parameter)
        {
            IsSyncRes = !IsSyncRes;
        }

        /// <summary>
        /// Проверка на возможность смены флага
        /// </summary>
        /// <param name="parameter"></param>
        /// <returns></returns>
        public bool CanExecuteSyncResCommand(object parameter)
        {
            if (IsListNullOrEmpty(_detailingProjects))
                return false;
            else
                return true;
        }

        RelayCommand _selectedDCommand;
        /// <summary>
        /// Синхронизировать только плановые ресурсы
        /// </summary>
        public ICommand SelectedDCommand
        {
            get
            {
                if (_selectedDCommand == null)
                    _selectedDCommand = new RelayCommand(ExecuteSelectedDCommand, CanExecuteSelectedDCommand);
                return _selectedDCommand;
            }
        }

        /// <summary>
        /// Смена флага на противоположное значение
        /// </summary>
        /// <param name="parameter"></param>
        public void ExecuteSelectedDCommand(object parameter)
        {
            System.Windows.Forms.MessageBox.Show(selectedDetailingProject.ToString());
        }

        /// <summary>
        /// Проверка на возможность смены флага
        /// </summary>
        /// <param name="parameter"></param>
        /// <returns></returns>
        public bool CanExecuteSelectedDCommand(object parameter)
        {
            if (!IsSyncRes || IsListNullOrEmpty(_detailingProjects))
            {
                IsSyncOnlyPlanRes = false;
                return false;
            }
            else
                return true;
        }


        RelayCommand _syncOnlyPlanResCommand;
        /// <summary>
        /// Синхронизировать только плановые ресурсы
        /// </summary>
        public ICommand SyncOnlyPlanResCommand
        {
            get
            {
                if (_syncOnlyPlanResCommand == null)
                    _syncOnlyPlanResCommand = new RelayCommand(ExecuteSyncOnlyPlanResCommand, CanExecuteSyncOnlyPlanResCommand);
                return _syncOnlyPlanResCommand;
            }
        }

        /// <summary>
        /// Смена флага на противоположное значение
        /// </summary>
        /// <param name="parameter"></param>
        public void ExecuteSyncOnlyPlanResCommand(object parameter)
        {
            IsSyncOnlyPlanRes = !IsSyncOnlyPlanRes;
        }

        /// <summary>
        /// Проверка на возможность смены флага
        /// </summary>
        /// <param name="parameter"></param>
        /// <returns></returns>
        public bool CanExecuteSyncOnlyPlanResCommand(object parameter)
        {
            if (!IsSyncRes || IsListNullOrEmpty(_detailingProjects))
            {
                IsSyncOnlyPlanRes = false;
                return false;
            }
            else
                return true;
        }

        #region Кнопка Отмена
        public Action CloseAction { get; set; }

        public bool CanClose { get; set; } = true;


        private RelayCommand _closeCommand;

        /// <summary>
        /// 
        /// </summary>
        public ICommand CloseCommand
        {
            get
            {
                if (_closeCommand == null)
                {
                    _closeCommand = new RelayCommand(param => Close(), param => CanClose);
                }
                return _closeCommand;
            }

        }

        public void Close()
        {
            CloseAction(); // Invoke the Action previously defined by the View
        }

        #endregion

        RelayCommand _syncCommand;
        /// <summary>
        /// Команда кнопки Синхронизировать
        /// </summary>
        public ICommand SyncCommand
        {
            get
            {
                if (_syncCommand == null)
                    _syncCommand = new RelayCommand(ExecuteSyncCommand, CanExecuteSyncCommand);
                return _syncCommand;
            }
        }


        /// <summary>
        /// Действие по нажатию кнопки
        /// </summary>
        /// <param name="parameter"></param>
        public void ExecuteSyncCommand(object parameter)
        {

        }

        /// <summary>
        /// Условия нажатия кнопки
        /// </summary>
        /// <param name="parameter"></param>
        public bool CanExecuteSyncCommand(object parameter)
        {
            if (IsListNullOrEmpty(_detailingProjects))
                return false;
            else
                return true;
        }

        RelayCommand _buildTreeCommand;
        /// <summary>
        /// команда кнопки "Загрузить"
        /// </summary>
        public ICommand BuildTreeCommand
        {
            get
            {
                if (_buildTreeCommand == null)
                    _buildTreeCommand = new RelayCommand(ExecuteBuildTreeCommand, CanExecuteBuildTreeCommand);
                return _buildTreeCommand;
            }
        }

        /// <summary>
        /// Действие кнопки "Загрузить"
        /// </summary>
        /// <param name="parameter"></param>
        public void ExecuteBuildTreeCommand(object parameter)
        {
            Synchronization.SynchronizingСomposition(startRefObject, SelectedDetailingProject as ReferenceObject, ref _tree);
            //Clients.Add(CurrentClient);
            //CurrentClient = null;
        }

        /// <summary>
        /// Условия выполнения кнопки "Загрузить"
        /// </summary>
        /// <param name="parameter"></param>
        /// <returns></returns>
        public bool CanExecuteBuildTreeCommand(object parameter)
        {

            if (IsListNullOrEmpty(_detailingProjects))
                return false;
            else
                return true;

        }

        #endregion

        /// <summary>
        /// Очистка коллекции
        /// </summary>
        protected override void OnDispose()
        {
            //this.Clients.Clear();
        }
    }
}
