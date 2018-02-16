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

        public MainWindowViewModel(ReferenceObject startRefObject)
        {
            this.startRefObject = startRefObject;

            if (startObject == null)
                startObject = Factory.Create_ProjectManagementWork(startRefObject);

            if (IsListNullOrEmpty(DetailingProjects))
                ShowError("Синхронизация состава работа", "Ошибка, выбранная детализация не найдена!");

            return;
        }


        public bool IsSyncRes
        {
            get { return ControlDialog.IsSyncRes; }
            set { OnPropertyChanged("IsSyncRes"); ControlDialog.IsSyncRes = value; }
        }



        public bool IsSyncOnlyPlanRes
        {
            get { return ControlDialog.IsSyncOnlyPlanRes; }
            set { OnPropertyChanged("IsSyncOnlyPlanRes"); ControlDialog.IsSyncOnlyPlanRes = value; }
        }

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
        ObservableCollection<ProjectManagementWork> _detailingProjects;
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


        private static List<ReferenceObject> зависимостиДетализации;

        public List<ReferenceObject> ЗависимостиДетализации
        {
            get
            {
                if (IsListNullOrEmpty(зависимостиДетализации))
                {
                    зависимостиДетализации = new List<ReferenceObject>();
                    зависимостиДетализации = Synchronization.GetDependenciesObjects(false, StartObject);
                }
                return зависимостиДетализации;
            }
        }

        void ShowError(string caption, string message)
        {
            System.Windows.Forms.MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1, System.Windows.Forms.MessageBoxOptions.ServiceNotification);
        }

        bool IsListNullOrEmpty<T>(IEnumerable<T> list)
        {

            if (list == null || list.Count() == 0)
                return true;

            return false;
        }


        public object Show { get; internal set; }

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


        RelayCommand _syncResCommand;
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
        /// Добавление в список
        /// </summary>
        /// <param name="parameter"></param>
        public void ExecuteSyncResCommand(object parameter)
        {
            IsSyncRes = !IsSyncRes;
        }

        /// <summary>
        /// Проверка перед добавлением
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


        RelayCommand _syncOnlyPlanResCommand;
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
        /// Добавление в список
        /// </summary>
        /// <param name="parameter"></param>
        public void ExecuteSyncOnlyPlanResCommand(object parameter)
        {

            IsSyncOnlyPlanRes = !IsSyncOnlyPlanRes;

            //Clients.Add(CurrentClient);
            //CurrentClient = null;
        }

        /// <summary>
        /// Проверка перед добавлением
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


        public Action CloseAction { get; set; }
        public bool CanClose { get; set; } = true;

        private RelayCommand _closeCommand;

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




        RelayCommand _syncCommand;
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
        /// Добавление в список
        /// </summary>
        /// <param name="parameter"></param>
        public void ExecuteSyncCommand(object parameter)
        {
            //Clients.Add(CurrentClient);
            //CurrentClient = null;
        }

        /// <summary>
        /// Добавление в список
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
        /// Добавление в список
        /// </summary>
        /// <param name="parameter"></param>
        public void ExecuteBuildTreeCommand(object parameter)
        {

            //Clients.Add(CurrentClient);
            //CurrentClient = null;
        }

        /// <summary>
        /// Проверка перед добавлением
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



        /// <summary>
        /// Очистка коллекции
        /// </summary>
        protected override void OnDispose()
        {
            //this.Clients.Clear();
        }
    }
}
