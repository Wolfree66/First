using System.Collections.ObjectModel;

namespace WpfApp_DialogueAddSignatories.Model
{
    public interface ITreeNode
    {
        ITreeNode Parent { get; }

        ObservableCollection<ITreeNode> Children { get; }
    }
}
