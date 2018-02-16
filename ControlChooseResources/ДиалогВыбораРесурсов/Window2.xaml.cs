using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ControlChooseResources
{
    /// <summary>
    /// Interaction logic for Window2.xaml
    /// </summary>
    public partial class Window2 : Window, INotifyPropertyChanged
    {
        public Window2()
        {
            FillTheStatuses();
            InitializeComponent();
            
            //start_startDate.SelectedDate = Start_StartDate;
            //start_startDate.DisplayDate = Start_StartDate;
        }

        void FillTheStatuses()
        {
            WorkStatuses = new List<WorkStatus>();
            string[] statuses = { "Данные не корректны.", "Выполнено", "Срок не наступил", "Не начиналось", "В работе", "Не выполнена" };
            foreach (string item in statuses)
            {
                WorkStatus ws = new WorkStatus();
                ws.IsChecked = true;
                ws.Status = item;
                WorkStatuses.Add(ws);
            }
            //this.checkedListBox1.CheckedItems = this.checkedListBox1.Items;
        }

        public List<WorkStatus> WorkStatuses { get; set; }

        public List<Resource> ChosenResources { get; set; }

        // свойство зависимостей
        public static readonly DependencyProperty Start_StartDateProperty =
          DependencyProperty.Register("Start_StartDate", typeof(DateTime), typeof(Window2), new FrameworkPropertyMetadata(null));
        public DateTime Start_StartDate
        {
            get { return (DateTime)GetValue(Start_StartDateProperty); }
            set
            {
                SetValue(Start_StartDateProperty, value);
                NotifyPropertyChanged("Start_StartDate");
            }
        }
        DateTime _Start_EndDate;
        public DateTime Start_EndDate
        {
            get
            {
                return this._Start_EndDate;
            }
            set
            {
                if (this._Start_EndDate != value)
                {
                    this._Start_EndDate = value;
                    OnPropertyChanged("Start_EndDate");
                }
            }
        }
        DateTime _End_StartDate;
        public DateTime End_StartDate {
            get
            {
                return this._End_StartDate;
            }
            set
            {
                if (this._End_StartDate != value)
                {
                    this._End_StartDate = value;
                    OnPropertyChanged("End_StartDate");
                }
            }
        }
        DateTime _End_EndDate;
        public DateTime End_EndDate {
            get
            {
                return this._End_EndDate;
            }
            set
            {
                if (this._End_EndDate != value)
                {
                    this._End_EndDate = value;
                    OnPropertyChanged("End_EndDate");
                }
            }
        }
        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            //MessageBox.Show(End_EndDate.ToString());
            //MessageBox.Show(end_endDate.SelectedDate.ToString());
            this.Close();
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            ChosenResources = uccr.CheckedResources;
            /*
            StringBuilder str = new StringBuilder();
            foreach (var item in result)
            {
                str.AppendLine(item.Parent.Name + " - " + item.Name);
            }
            ChosenResources = result;
            MessageBox.Show(str.ToString());
            */
            //ChosenResources = result;
            this.DialogResult = true;
            //this.Close();
        }
        #region 
        public event PropertyChangedEventHandler PropertyChanged;

        // This method is called by the Set accessor of each property.
        // The CallerMemberName attribute that is applied to the optional propertyName
        // parameter causes the property name of the caller to be substituted as an argument.
        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        protected virtual void OnPropertyChanged(String propertyName)
        {
            if (System.String.IsNullOrEmpty(propertyName))
            {
                return;
            }
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        #endregion

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //this.SizeToContent = SizeToContent.Width;
            //Start_StartDate = DateTime.Today.AddDays(-5);
            //End_StartDate = DateTime.Today.AddDays(-5);
            /*start_startDate.SelectedDate = Start_StartDate;
            start_startDate.DisplayDate = Start_StartDate;
            end_startDate.UpdateLayout();
            */
        }

        private void Expander_Expanded(object sender, RoutedEventArgs e)
        {
            this.SizeToContent = SizeToContent.Width;
        }
    }

}
