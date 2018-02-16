using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using TFlex.DOCs.Model;

namespace ControlChooseResources
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
           // MessageBox.Show(DateTime.Now.CompareTo(DateTime.Now.AddDays(1)).ToString());
            //InitializeComponent();
            ServerGateway.Connect("tf-test");
            if (!ServerGateway.Connect(false))
            {
                Close();
                return;
            }
            else { }
            Window2 w1 = new Window2();
            w1.ShowDialog();
            Close();
        }
    }
}
