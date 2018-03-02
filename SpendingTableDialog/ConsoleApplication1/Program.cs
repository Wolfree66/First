using SpendingTableDialog;
using SpendingTableDialog.Models;
using System.Windows;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;

namespace ConsoleApplication1
{
    
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            List<string> stringValues = new List<string> {"карта", "счёт" };
            List<TableItem> list = new List<ConsoleApplication1.TableItem> {
                new TableItem("fgfg", 1000, "карта", stringValues),
                new TableItem("fваавп вап ва пва g", 2000, "счёт", stringValues),
                new TableItem("fвапвапвапвg", 3000, "счёт", stringValues)
            };

            SpendingTableView window = new SpendingTableView(list);
            window.ShowDialog();
            foreach (var item in list)
            {
                Console.WriteLine(item.Item + "\t" + item.DoubleValue + "\t" + item.StringValue);
            }
            Console.ReadKey();
        }
    }
    class TableItem : ITableItem
    {
        public TableItem(object item, double doubleValue, string stringValue, IReadOnlyCollection<string> stringValues)
        {
            this.Item = item;
            this.DoubleValue = doubleValue;
            this.StringValue = stringValue;
            this.StringValues = stringValues;
        }
        public double DoubleValue
        {
            get;set;
        }

        public object Item
        {
            get; set;
        }

        public string StringValue
        {
            get; set;
        }

        public IReadOnlyCollection<string> StringValues
        {
            get; set;
        }
    }
}
