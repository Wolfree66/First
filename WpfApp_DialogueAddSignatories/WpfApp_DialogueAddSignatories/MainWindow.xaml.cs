using System;
using System.Windows;
using System.Windows.Controls;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using TFlex.DOCs.Model;

namespace WpfApp_DialogueAddSignatories
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Presenter Presenter;
        public MainWindow(ServerConnection connection, TFlex.DOCs.Model.References.ReferenceObject currentTechnicalSolution, string nameServer = "TFLEX")
        {
            InitializeComponent();
            InitializeListBox();
            if (Presenter == null)
                Presenter = new Presenter(connection, this, currentTechnicalSolution, nameServer);
        }

        private void InitializeListBox()
        {
            System.Collections.Generic.List<string> posts = new System.Collections.Generic.List<string>();
            posts.Add("Главный конструктор по техническим средствам обучения");
            posts.Add("Главный конструктор по беспилотным системам");
            this.listBox_posts.ItemsSource = posts;
        }

        public event EventHandler FindEmployeeAtPostEvent = null;
        public event EventHandler SelectEmployeeEvent = null;
        public event EventHandler GetAllUserEvent = null;
        public event EventHandler OKEvent = null;
        public event EventHandler SelectProgramManagerEvent = null;
        public event EventHandler ClearComboBoxEvent = null;
        private void button_accept_Click(object sender, RoutedEventArgs e)
        {
            OKEvent.Invoke(sender, e);
        }

        private void listBox_posts_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FindEmployeeAtPostEvent.Invoke(sender, e);
        }


        private void comboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SelectProgramManagerEvent.Invoke(sender, e);
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            GetAllUserEvent.Invoke(sender, e);
        }

        private void Window_AccessKeyPressed(object sender, System.Windows.Input.AccessKeyPressedEventArgs e)
        {

        }

        private void button_clean_Click(object sender, RoutedEventArgs e)
        {
            ClearComboBoxEvent.Invoke(sender, e);
        }

        private void listBox_users_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SelectEmployeeEvent.Invoke(sender, e);
        }
    }
}