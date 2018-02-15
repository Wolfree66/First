using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model.References;

namespace TFLEXDocsClasses
{
    class CancelariaRegistryJournal
    {
        public CancelariaRegistryJournal(ReferenceObject referenceObject)
        {
            this.ReferenceObject = referenceObject;
        }

        public ReferenceObject ReferenceObject { get; private set; }

        string _Name;
        public string Name
        {

            get
            {
                if (_Name == null)
                {
                    Guid name_Guid = new Guid("61adbc9b-72f4-415e-a144-9832828cd701");
                    _Name = this.ReferenceObject[name_Guid].GetString();
                }
                return _Name;
            }
        }
    }
}
