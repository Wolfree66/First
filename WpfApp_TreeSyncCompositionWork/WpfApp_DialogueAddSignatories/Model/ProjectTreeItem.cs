using System.Collections.ObjectModel;
using TFlex.DOCs.Model.References;

namespace WpfApp_DialogueAddSignatories.Model
{
    public class ProjectTreeItem 
    {

        public ReferenceObject ReferenceObject;
        private string name;

        public string Name
        {
            get { return ReferenceObject.ToString(); }
            set { name = value; }
        }


        public ProjectTreeItem()
        {

        }
        public ProjectTreeItem(ReferenceObject referenceObject, bool IsForAdd)
        {
            this.ReferenceObject = referenceObject;
            this.IsForAdd = IsForAdd;
        }

        public bool IsForAdd { get; set; }


        ProjectTreeItem _Parent;
        public ProjectTreeItem Parent
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

        ObservableCollection<ProjectTreeItem> _Children;
        public ObservableCollection<ProjectTreeItem> Children
        {
            get
            {
                if (_Children == null)
                {
                    ObservableCollection<ProjectTreeItem> temp = new ObservableCollection<ProjectTreeItem>();
                    foreach (var child in this.ReferenceObject.Children)
                    {
                        ProjectTreeItem node = Factory.CreateProjectTreeItem(child);
                        if (node != null) temp.Add(node);
                    }
                    _Children = temp;
                }
                return _Children;
            }
            set
            {

            }
        }


    }
}
