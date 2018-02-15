using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Users;

namespace TFLEXDocsClasses
{
    /// <summary>
    /// Для работы с регистрационно-контрольными карточками Канцелярии
    /// ver 1.3
    /// </summary>
    class RegistryControlCard
    {

        bool ЕстьИзмененияДляСохраненияВБД;
        public RegistryControlCard(ReferenceObject original)
        {
            ReferenceObject = original;
            ЕстьИзмененияДляСохраненияВБД = false;
        }
        public RegistryControlCard()
        {
            ReferenceObject = null;
            ЕстьИзмененияДляСохраненияВБД = false;
        }

        /*public void RegistryControlCardForOfficialNote(OfficialNote officialNote)
            {
                ОбъектСправочника_ro = FindEarlierCreatedRCC(officialNote);
                if (ОбъектСправочника_ro == null)
                {
                    ОбъектСправочника_ro = CreateNewRCC_ro();
                    ЕстьИзмененияДляСохраненияВБД = true;
                }
                else ОбъектСправочника_ro.BeginChanges();
                AddDocument(officialNote.ReferenceObject);
                AddDocument(officialNote.GetReportFile());
                foreach (var item in officialNote.AdditionalFiles)
                {
                    AddDocument(item);

                }
                ОбъектСправочника_ro[RCC_param_FromDate_GUID].Value = officialNote.RegistryDate;
                RegistryNumber = officialNote.GetNumber();
                Performer = officialNote.Initiator; //ркк.СвязанныйОбъект["Исполнитель"] = Служебная.СвязанныйОбъект["Подчиненный"];
                WhoSigned = officialNote.GetApprover().ToString(); //ркк.Параметр["Подписал"] = Служебная.Параметр["Руководитель"];

                ОбъектСправочника_ro[RCC_param_Contents_GUID].Value = officialNote.Theme;//ркк.Параметр["Содержание"] = Служебная.Параметр["Тематика"];

                ПрименитьИзменения();
            }
         */
        private void AddDocument(ReferenceObject officialNote_ro)
        {
            ReferenceObject.AddLinkedObject(RCC_link_Documents_GUID, officialNote_ro);
            _Documents.Add(officialNote_ro);
        }

        /* private ReferenceObject FindEarlierCreatedRCC(OfficialNote officialNote)
             {
                 return RegistryControlCard_reference.Find(RegistryControlCard_reference.ParameterGroup[RCC_param_RegistryNumber_GUID], officialNote.GetNumber()).FirstOrDefault();
             }*/

        private ReferenceObject CreateNewRCC_ro()
        {
            return RegistryControlCard_reference.CreateReferenceObject(RegistryControlCard_Internal_Class);
        }

        public bool HasBibleORDItem()
        {
            ReferenceObject.ApplyChanges();
            ReferenceObject.Reload();
            if (ReferenceObject.GetObject(BORD_link_RegistryCancelariaCard_GUID) != null) return true;
            else return false;
        }

        public IEnumerable<ReferenceObject> GetAllLinkedFiles()
        {
            FileType fileType = (new FileReference(ReferenceObject.Reference.Connection)).Classes.FileBase;
            return ReferenceObject.GetObjects(RCC_link_Documents_GUID).Where(d => d.Class.IsInherit(fileType));
        }

        public int GetAccessLevel()
        {
            return ReferenceObject[RCC_param_AccessLevel_GUID].GetInt16();
        }

        List<ReferenceObject> _Documents;
        /// <summary>
        /// Документы
        /// </summary>
        public IEnumerable<ReferenceObject> Documents
        {
            get
            {
                if (_Documents == null) _Documents = ReferenceObject.GetObjects(RCC_link_Documents_GUID);
                return _Documents;
            }
        }

        ReferenceObject _Checker;
        /// <summary>
        /// Контролёр
        /// </summary>
        public ReferenceObject Checker
        {
            get
            {
                if (_Checker == null) _Checker = ReferenceObject.GetObject(RCC_link_Checker_GUID);
                return _Checker;
            }
            set
            {
                if (value == _Checker) return;
                _Checker = value;
                if (_Checker != null)
                {
                    ЕстьИзмененияДляСохраненияВБД = true;
                }
                ReferenceObject.SetLinkedObject(RCC_link_Checker_GUID, _Checker);
            }
        }

        ReferenceObject _Responsible;
        /// <summary>
        /// Ответственный
        /// </summary>
        public ReferenceObject Responsible
        {
            get
            {
                if (_Responsible == null) _Responsible = ReferenceObject.GetObject(RCC_link_Responsible_GUID);
                return _Responsible;
            }
            set
            {
                if (value == _Responsible) return;
                _Responsible = value;
                if (_Responsible != null)
                {
                    ЕстьИзмененияДляСохраненияВБД = true;
                }
                ReferenceObject.SetLinkedObject(RCC_link_Responsible_GUID, _Responsible);
            }
        }

        ReferenceObject _Performer;
        /// <summary>
        /// Исполнитель
        /// </summary>
        public ReferenceObject Performer
        {
            get
            {
                if (_Performer == null) _Performer = ReferenceObject.GetObject(RCC_link_Performer_GUID);
                return _Performer;
            }
            set
            {
                if (value == _Performer) return;
                _Performer = value;
                if (_Performer != null)
                {
                    ЕстьИзмененияДляСохраненияВБД = true;
                }
                ReferenceObject.SetLinkedObject(RCC_link_Performer_GUID, _Performer);
            }
        }
        string _WhoSigned;

        public string WhoSigned
        {
            get
            {
                if (_WhoSigned == null)
                    _WhoSigned = ReferenceObject[RCC_param_WhoSigned_GUID].GetString();
                return _WhoSigned;
            }
            set
            {
                if (value == _WhoSigned) return;
                _WhoSigned = value;
                if (_WhoSigned != null)
                {
                    ЕстьИзмененияДляСохраненияВБД = true;
                }
                ReferenceObject[RCC_param_WhoSigned_GUID].Value = _WhoSigned;
            }
        }

        string _Content;
        /// <summary>
        /// Возвращает содержание карточки
        /// </summary>
        public string Content
        {
            get
            {
                if (_Content == null) _Content = ReferenceObject[RCC_param_Content_GUID].GetString();
                return _Content;
            }
        }

        string _StorageFolder;
        /// <summary>
        /// возвращает папку хранения документа
        /// </summary>
        public string StorageFolder
        {
            get
            {
                if (_StorageFolder == null) _StorageFolder = ReferenceObject[RCC_param_StorageFolderName_GUID].GetString();
                return _StorageFolder;
            }
        }

        string _RegistryNumber;
        public string RegistryNumber
        {
            get
            {
                if (_RegistryNumber == null)
                    _RegistryNumber = ReferenceObject[RCC_param_RegistryNumber_GUID].GetString();
                return _RegistryNumber;
            }
            set
            {
                if (value == _RegistryNumber) return;
                if (_RegistryNumber != null)
                {
                    ЕстьИзмененияДляСохраненияВБД = true;
                }
                _RegistryNumber = value;
                ReferenceObject[RCC_param_RegistryNumber_GUID].Value = _RegistryNumber;
            }
        }

        DateTime? _RegistryDate;
        /// <summary>
        /// Возвращает дату "От"
        /// </summary>
        public DateTime? RegistryDate
        {
            get
            {
                if (_RegistryDate == null)
                {
                    DateTime date = this.ReferenceObject[RCC_param_FromDate_GUID].GetDateTime();
                    if (date != DateTime.MinValue) _RegistryDate = date;
                }
                return _RegistryDate;
            }
        }

        CancelariaRegistryJournal _RegistryJournal;
        public CancelariaRegistryJournal RegistryJournal
        {
            get
            {
                if (_RegistryJournal == null)
                {
                    ReferenceObject ro = ReferenceObject.GetObject(RCC_link_RegistryJournal_GUID);
                    if (ro != null) _RegistryJournal = new CancelariaRegistryJournal(ro);
                }
                return _RegistryJournal;
            }
            //set
            //{
            //    if (value == _RegistryJournal) return;
            //    if (_RegistryJournal != null)
            //    {
            //        ЕстьИзмененияДляСохраненияВБД = true;
            //    }
            //    _RegistryJournal = value;
            //    ReferenceObject.SetLinkedObject(RCC_link_RegistryJournal_GUID, _RegistryJournal);
            //}
        }

        /// <summary>
        /// Организации Адресаты исходящей корреспонденции
        /// </summary>
        /// <returns></returns>
        public IEnumerable<Organization> Get_ToOrganizations()
        {
            List<Organization> result = new List<Organization>();
            foreach (var item in this.ReferenceObject.GetObjects(RCC_link_Organizations_GUID))
            {
                result.Add(new Organization(item));
            }
            if (result.Count == 0)
            {
                foreach (var item in this.ReferenceObject.GetObjects(RCC_link_Organizations_GUID_Obsolete))
                {
                    result.Add(new Organization(item));
                }
            }

            return result;
        }

        string _From;
        public string From
        {
            get
            {
                if (_From == null)
                    _From = this.ReferenceObject[RCC_param_From_GUID].GetString();

                return _From;
            }
        }

        List<ReferenceObject> _SignedTo;

        /// <summary>
        /// список Подписано/Кому
        /// </summary>
        public IEnumerable<ReferenceObject> SignedTo
        {
            get
            {
                if (_SignedTo == null)
                    _SignedTo = ReferenceObject.GetObjects(RCC_link_SignedTo_GUID);
                else _SignedTo = new List<ReferenceObject>();

                return _SignedTo;
            }
        }

        /// <summary>
        /// Все пользователи имеющие доступ (участники + доп.доступ)
        /// </summary>
        /// <returns></returns>
        public IEnumerable<ReferenceObject> GetUsersForAccess()
        {
            List<ReferenceObject> result = new List<ReferenceObject>();

            result.AddRange(this.SignedTo);
            ReferenceObject performer = ReferenceObject.GetObject(RCC_link_Performer_GUID);
            if (performer != null)
                result.Add(performer);
            result.AddRange(ReferenceObject.GetObjects(RCC_link_ExtraAccess_GUID));
            return result.Distinct().ToList();
        }
        List<ReferenceObject> СписокЛюдейИзДополнительногоДоступа()
        {
            return ReferenceObject.GetObjects(RCC_link_ExtraAccess_GUID);
        }
        public void ДобавитьВДополнительныйДоступ(ReferenceObject пользователь)
        {
            if (пользователь == null) return;
            if ((пользователь as User) == null) throw new Exception(пользователь.ToString() + " не является пользователем");
            if (СписокЛюдейИзДополнительногоДоступа().Contains(пользователь)) return;
            ((ReferenceObject)ReferenceObject).AddLinkedObject(RCC_link_ExtraAccess_GUID, пользователь);
            ЕстьИзмененияДляСохраненияВБД = true;
            return;
        }

        public bool ПрименитьИзменения()
        {
            if (ЕстьИзмененияДляСохраненияВБД) return ReferenceObject.ApplyChanges();
            return false;
        }

        public ReferenceObject ReferenceObject;

        #region Справочники
        private static Guid reference_RCC_Guid = new Guid("80831dc7-9740-437c-b461-8151c9f85e59");    //Guid справочника - "Регистрационно-контрольные карточки"
        private static Guid RCC_Internal_class_Guid = new Guid("271f689e-3986-40d5-ae0c-c56d74bf8f3d");    //Guid типа "Внутренняя" - справочника "Регистрационно-контрольные карточки"
        private static Guid RCC_link_ExtraAccess_GUID = new Guid("7f823151-f3ca-40be-97b0-a1c69126a027"); //связь со справочником Группы и пользователи - Дополнительный доступ
        private static Guid RCC_link_SignedTo_GUID = new Guid("a3746fbd-7bda-4326-8e4b-f6bbefc2ecce"); //связь N:N со Подписал/Кому
        private static Guid RCC_link_Responsible_GUID = new Guid("6a69454d-7803-4cd8-8874-cde190185b7b"); //связь N:1  Ответственный
        private static Guid RCC_link_Performer_GUID = new Guid("17330c2c-589f-4838-b6d1-8a9214666c4e"); //связь N:1  Исполнитель
        private static Guid RCC_link_Checker_GUID = new Guid("cda6f4f3-0c03-43ad-819c-b89cfca1579f"); //связь N:1  Контроллер
        private static Guid RCC_link_Documents_GUID = new Guid("aab1cd65-fe2d-4fc0-8db2-4ff2b3d215d5"); //связь N:N на любой справочник Документы
        public static Guid RCC_link_Organizations_GUID_Obsolete = new Guid("4840f02e-19b1-41ff-a063-b98dfa7639da"); //связь N:N со справочником Организации
        public static Guid RCC_link_Organizations_GUID = new Guid("c6afa49d-5c14-4f36-a488-6bce169c1a31"); //связь N:N со справочником Данные организаций
        public static Guid RCC_link_RegistryJournal_GUID = new Guid("2158039e-d86d-44af-b384-761bc7f233fb"); //связь N:1 со справочником Журнал регистрации

        private static Guid RCC_param_AccessLevel_GUID = new Guid("03209f40-7904-4b07-9843-c483050cc1fa"); //параметр "Уровень доступа"
        private static Guid RCC_param_Content_GUID = new Guid("b85553c9-2e2d-4f27-96fa-9ebb8e365747"); //параметр "Содержание"
        private static Guid RCC_param_FromDate_GUID = new Guid("9930335a-8703-47a1-8dc8-49e17a959d20"); //параметр "От"
        private static Guid RCC_param_From_GUID = new Guid("98ac9043-9e24-4948-b853-a3b93a64cb8e"); //параметр "Откуда поступил"

        private static Guid RCC_param_RegistryNumber_GUID = new Guid("ac452c8a-b17a-4753-be3f-195d76480d52"); //параметр "Регистрационный номер"
        private static Guid RCC_param_WhoSigned_GUID = new Guid("87b53d7d-be4e-4eff-9fca-511be18d2498"); //параметр "Подписал"
        private static Guid RCC_param_StorageFolderName_GUID = new Guid("3073a808-4e9a-4a55-97e5-a793f32fdc5f"); //параметр "Папка хранения документа"


        private static Guid BORD_link_RegistryCancelariaCard_GUID = new Guid("776b674c-4944-460a-a8e0-729efa58911e"); //cвязь 1:1 Регистрационно-контрольная карта Справочника Канцелярия

        //поля
        private Reference _RegistryControlCardReference;
        private ReferenceInfo _RCCReferenceInfo;
        //Типы
        private static ClassObject _RegistryControlCard_Internal_class;
        private Reference GetReference(ref Reference reference, ReferenceInfo referenceInfo)
        {
            if (reference == null)
                reference = referenceInfo.CreateReference();


            return reference;
        }


        /// <summary>
        /// Справочник "Регистрационно-контрольные карточки"
        /// </summary>
        private Reference RegistryControlCard_reference
        {
            get
            {
                if (_RegistryControlCardReference == null)

                    return GetReference(ref _RegistryControlCardReference, RCCReferenceInfo);

                return _RegistryControlCardReference;
            }
        }



        private ReferenceInfo RCCReferenceInfo
        {
            get { return GetReferenceInfo(ref _RCCReferenceInfo, reference_RCC_Guid); }
        }

        private ReferenceInfo GetReferenceInfo(ref ReferenceInfo referenceInfo, Guid referenceGuid)
        {
            if (referenceInfo == null)
                referenceInfo = Connection.ReferenceCatalog.Find(referenceGuid);

            return referenceInfo;
        }
        /// <summary>
        /// Тип - Регистрационно-контрольная карточка
        /// </summary>
        private ClassObject RegistryControlCard_Internal_Class
        {
            get
            {
                if (_RegistryControlCard_Internal_class == null)
                    _RegistryControlCard_Internal_class = RegistryControlCard_reference.Classes.Find(RCC_Internal_class_Guid);

                return _RegistryControlCard_Internal_class;
            }
        }
        ServerConnection Connection { get { return ServerGateway.Connection; } }

        public object Type
        {
            get
            {
                return this.ReferenceObject.Class.Name;
            }
        }




        #endregion
    }
}
