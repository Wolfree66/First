using SpendingTableDialog.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;

namespace SpendingTableDialog.ViewModels
{
    public class TableItemViewModel : DependencyObject, ITableItem, INotifyPropertyChanged
    {
        public TableItemViewModel(ITableItem item)
        {
            this.Item = item;
        }
        public string ExpenseName { get; set; }

        public static readonly DependencyProperty DoubleValueProperty =
       DependencyProperty.Register("DoubleValue", typeof(double), typeof(TableItemViewModel), new FrameworkPropertyMetadata(0.0, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, new PropertyChangedCallback(OnDoubleValuePropertyChanged)/*, new CoerceValueCallback(NumHoursCoerceValue)*/)/*, new ValidateValueCallback (IsNumHoursValid)*/);

        private static void OnDoubleValuePropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d == null) return;
            ((d as TableItemViewModel).Item as ITableItem).DoubleValue = (double)e.NewValue;
            //throw new NotImplementedException();
        }

        public double DoubleValue
        {
            get
            {
                return (double)GetValue(DoubleValueProperty);
            }
            set
            {
                SetValue(DoubleValueProperty, (double)value);
                //NotifyPropertyChanged("NumHours");
            }
        }

        public IReadOnlyCollection<string> StringValues { get; set; }

        public static readonly DependencyProperty StringValueProperty =
       DependencyProperty.Register("StringValue", typeof(string), typeof(TableItemViewModel), new FrameworkPropertyMetadata("", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, new PropertyChangedCallback(OnStringValuePropertyChanged)/*, new CoerceValueCallback(NumHoursCoerceValue)*/)/*, new ValidateValueCallback (IsNumHoursValid)*/);

        private static void OnStringValuePropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d == null) return;
            ((d as TableItemViewModel).Item as ITableItem).StringValue = (string)e.NewValue;
        }

        public string StringValue
        {
            get
            {
                return (string)GetValue(StringValueProperty);
            }
            set
            {
                SetValue(StringValueProperty, (string)value);
                //NotifyPropertyChanged("NumHours");
            }
        }

        public object Item
        {
            get;
            set;
        }

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
