using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFLEXDocsClasses;

namespace PMTreeSync.ViewModels
{
    class WorkTreeItemViewModel
    {
        public WorkTreeItemViewModel(ProjectManagementWork work)
        {
            this.Item = work;
        }
        public ProjectManagementWork Item { get; set; }

        public bool IsSelected { get; set; }

        public WorkTreeItemViewModel Parent { get; set; }

        public bool IsVisible { get; set; }

        public override string ToString()
        {
            return this.Item.ToString();
        }


        public ObservableCollection<WorkTreeItemViewModel> Children
        {
            get
            {
                ObservableCollection<WorkTreeItemViewModel> children = new ObservableCollection<WorkTreeItemViewModel>();
                foreach (var item in this.Item.Children)
                {
                    children.Add(new WorkTreeItemViewModel(item));
                }
                return children;
            }
        }
    }
}
