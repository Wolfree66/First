using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.References.ProjectManagement;

namespace WpfApp_DialogueAddSignatories.Model
{
    public class Node
    {
        public bool IsForAdd { get; set; }
        public Node(ProjectElement pe, bool IsForAdd)
        {
            PE = pe;
            this.IsForAdd = IsForAdd;
        }
        public ProjectElement PE { get; set; }
        public ObservableCollection<Node> Children { get; set; }
        private string name;

        public string Name
        {
            get { return this.PE.Name; }

        }

        public override string ToString()
        {
            return this.PE.Name;
        }

    }
}
