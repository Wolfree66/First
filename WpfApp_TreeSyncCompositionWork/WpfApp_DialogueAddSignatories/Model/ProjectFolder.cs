using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model.References;

namespace WpfApp_DialogueAddSignatories.Model
{
    public class ProjectFolder : ProjectTreeItem
    {
        public ProjectFolder(ReferenceObject referenceObject)
        {
            this.ReferenceObject = referenceObject;
        }
        public string Name { get { return this.ReferenceObject.ToString(); } }
        //public bool IsNull { get; private set; }

        public static Guid PM_class_ProjectFolder_Guid = new Guid("ae994d8a-7193-4ee3-8ef6-e443b23f7f54"); //тип - "Папка"
    }
}
