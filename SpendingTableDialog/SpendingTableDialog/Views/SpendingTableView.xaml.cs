using SpendingTableDialog.Commands;
using SpendingTableDialog.Models;
using SpendingTableDialog.ViewModels;
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

namespace SpendingTableDialog
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class SpendingTableView : Window
    {
        public SpendingTableView(IReadOnlyCollection<ITableItem> listExpenses)
        {
            InitializeComponent();
            viewModel = new SpendingTableViewModel(listExpenses);
            this.DataContext = viewModel;
        }

        SpendingTableViewModel viewModel;
        private RelayCommand _OKButtonCommand;

        public ICommand OKButtonCommand
        {
            get
            {
                if (_OKButtonCommand == null)
                {
                    _OKButtonCommand = new RelayCommand(this.OKButtonClick, this.CanExecuteCommandOKButton);
                }
                return _OKButtonCommand;
            }
        }
        private void OKButtonClick(object param)
        {
            this.DialogResult = true;
            //MessageBox.Show("Click");
        }

        private bool CanExecuteCommandOKButton(object param)
        {
            return true;
        }
    }
}
