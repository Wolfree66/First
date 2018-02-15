using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Users;

namespace TFLEXDocsClasses
{
    public class Organization : IMailRecipient
    {
        public Organization(ReferenceObject referenceObject)
        {
            this.ReferenceObject = referenceObject;
        }

        ReferenceObject ReferenceObject;

        string _Email;
        public string Email
        {
            get
            {
                if (_Email == null)
                {
                    Guid emailGuid = new Guid("0f4a1439-7f2c-495d-940a-bb41937b0c23");
                    string emailAddress = this.ReferenceObject[emailGuid].GetString();
                    if (!string.IsNullOrWhiteSpace(emailAddress)) _Email = emailAddress;
                }
                return _Email;
            }
        }

        string _Name;
        public string Name
        {
            get
            {
                if (_Name == null)
                {
                    Guid nameGuid = new Guid("e434c150-94e4-4cb6-a1de-9fc8b6408380");
                    _Name = this.ReferenceObject[nameGuid].GetString();
                }
                return _Name;
            }
        }

        public UserReferenceObject UserReferenceObject
        {
            get
            {
                return null;
            }
        }
    }
}
