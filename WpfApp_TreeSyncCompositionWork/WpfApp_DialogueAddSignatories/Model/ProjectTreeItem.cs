using System.Collections.ObjectModel;
using TFlex.DOCs.Model.References;

namespace WpfApp_DialogueAddSignatories.Model
{
    public class ProjectTreeItem : ITreeNode
    {


        //public ProjectTreeItem(ReferenceObject referenceObject, bool IsForAdd)
        //{
        //    this.ReferenceObject = referenceObject;
        //    this.IsForAdd = IsForAdd;
        //}

        public bool IsForAdd { get; set; }

        public ReferenceObject ReferenceObject;

        ITreeNode _Parent;
        public ITreeNode Parent
        {
            get
            {
                if (_Parent == null)
                {
                    ReferenceObject parent = this.ReferenceObject.Parent;
                    if (parent != null) _Parent = Factory.CreateProjectTreeItem(parent);
                }
                return _Parent;
            }
        }

        ObservableCollection<ITreeNode> _Children;
        public ObservableCollection<ITreeNode> Children
        {
            get
            {
                if (_Children == null)
                {
                    ObservableCollection<ITreeNode> temp = new ObservableCollection<ITreeNode>();
                    foreach (var item in this.ReferenceObject.Children)
                    {
                        ITreeNode node = Factory.CreateProjectTreeItem(item);
                        if (node != null) temp.Add(node);
                    }
                    _Children = temp;
                }
                return _Children;
            }
        }
    }
}
