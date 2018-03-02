using SpendingTableDialog.Commands;
using SpendingTableDialog.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;

namespace SpendingTableDialog.ViewModels
{
    public class SpendingTableViewModel : INotifyPropertyChanged
    {

        public SpendingTableViewModel(IReadOnlyCollection<ITableItem> listExpenses)
        {
            this.Items = new ObservableCollection<ViewModels.TableItemViewModel>();
            foreach (var expense in listExpenses)
            {
                this.Items.Add(CreateNewTableItem(expense));
            }
        }

        private TableItemViewModel CreateNewTableItem(ITableItem expense)
        {
            TableItemViewModel result = new ViewModels.TableItemViewModel(expense);
            result.DoubleValue = expense.DoubleValue;
            result.ExpenseName = expense.Item.ToString();
            result.StringValue = expense.StringValue;
            result.StringValues = expense.StringValues;
            return result;
        }

        public ObservableCollection<TableItemViewModel> Items { get; set; }
        

        public event PropertyChangedEventHandler PropertyChanged;
        public event EventHandler CanExecuteChanged;

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
    }
}
