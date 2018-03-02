using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpendingTableDialog.Models
{
    public interface ITableItem
    {
        object Item { get; }
        string StringValue { get; set; }
        double DoubleValue { get; set; }
        IReadOnlyCollection<string> StringValues { get; }
    }
}
