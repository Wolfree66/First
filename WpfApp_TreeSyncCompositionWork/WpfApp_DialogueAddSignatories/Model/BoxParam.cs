using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.References.ProjectManagement;

namespace WpfApp_DialogueAddSignatories.Model
{
    public struct BoxParam
    {
        public ProjectManagementWork currentObject { get; set; }
        public ProjectManagementWork Detailing { get; set; }
        public bool IsCopyRes { get; set; }
        public bool IsCopyPlan { get; set; }

    }
}
