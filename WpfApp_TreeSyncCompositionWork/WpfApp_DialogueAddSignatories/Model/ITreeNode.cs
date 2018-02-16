using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WpfApp_DialogueAddSignatories.Model;

namespace WpfApp_DialogueAddSignatories.Model
{
    public interface ITreeNode
    {
        ITreeNode Parent { get; }

        ObservableCollection<ITreeNode> Children { get; }
    }
}
