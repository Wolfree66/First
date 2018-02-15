using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Users;

namespace TFLEXDocsClasses
{
    /// <summary>
    /// Запись в Журнале редактирования планов
    /// ver. 1.3
    /// </summary>
    public class EditReasonNote
    {
        internal EditReasonNote(ProjectManagementWork work, string editType, string reason, bool needManagerAction = false)
        {
            //#if test
            //            MessageBox.Show("referenceInfo - " + EditReasonNote.ReferenceInfo.ToString());
            //            MessageBox.Show("Reference - " + Reference.ToString());
            //            MessageBox.Show("classObject - " + classObject.ToString());
            //#endif
            ReferenceObject newObject = Reference.CreateReferenceObject(classObject);// создание

            // Параметр "Тип редактирования"
            int editTypeCode = Array.IndexOf(TypeNames, editType);
            newObject[EditReasonNote.ERN_param_EditType_Guid].Value = editTypeCode;
            newObject[EditReasonNote.ERN_param_Reason_Guid].Value = reason;
            newObject[EditReasonNote.ERN_param_NeedActionManager_Guid].Value = needManagerAction;
            newObject.SetLinkedObject(EditReasonNote.ERN_link_ToPrjctMgmntN1_Guid, work.ReferenceObject);
            newObject.EndChanges();
            this.ReferenceObject = newObject;
        }
        public EditReasonNote(ReferenceObject changeReason_ro)
        {
            ReferenceObject = changeReason_ro;
        }

        public EditReasonNote(Guid changeReasonGuid)
        {
            ReferenceObject search = EditReasonNote.Reference.Find(changeReasonGuid);
            if (search == null) throw new ArgumentException("Запись в журнале редактирования, с Guid = " + changeReasonGuid.ToString() + ", не найдена!");
            this.ReferenceObject = search;
        }

        ReferenceObject ReferenceObject;

        /// <summary>
        /// Автор записи
        /// </summary>
        public User Author
        {
            get
            {
                return this.ReferenceObject.SystemFields.Author;
            }
        }

        /// <summary>
        /// Guid объекта в БД
        /// </summary>
        public Guid Guid
        {
            get
            {
                if (ReferenceObject != null) return ReferenceObject.SystemFields.Guid;
                else return Guid.Empty;
            }
        }

        string _Reason;
        /// <summary>
        /// Причина редактирования
        /// </summary>
        public string Reason
        {
            get
            {
                if (_Reason == null)
                    _Reason = ReferenceObject[EditReasonNote.ERN_param_Reason_Guid].GetString();
                return _Reason;
            }
        }

        int? _TypeCode;
        /// <summary>
        /// Код Типа редактирования
        /// </summary>
        public int TypeCode
        {
            get
            {
                if (_TypeCode == null)
                    _TypeCode = ReferenceObject[EditReasonNote.ERN_param_EditType_Guid].GetInt16();
                return (int)_TypeCode;
            }
        }

        ProjectManagementWork _Work;
        /// <summary>
        /// Связанная работа
        /// </summary>
        internal ProjectManagementWork Work
        {
            get
            {
                if (_Work == null && this.ReferenceObject != null)
                {
                    ReferenceObject work_ro = this.ReferenceObject.GetObject(ERN_link_ToPrjctMgmntN1_Guid);
                    _Work = new ProjectManagementWork(work_ro);
                }
                return _Work;
            }
        }

        bool? _IsActionOfManagerNeeded;
        public bool IsActionOfManagerNeeded
        {
            get
            {
                if (_IsActionOfManagerNeeded == null)
                {
                    _IsActionOfManagerNeeded = false;
                    if (this.ReferenceObject != null)
                    {
                        _IsActionOfManagerNeeded = this.ReferenceObject[ERN_param_NeedActionManager_Guid].GetBoolean();
                    }
                }
                return (bool)_IsActionOfManagerNeeded;
            }
        }

        /// <summary>
        /// Тип редактирования
        /// </summary>
        public string TypeName
        { get { return TypeNames[TypeCode]; } }

        /*public static Dictionary<int, string> TypeStrings = new Dictionary<int, string>
        {
            { (int)Types.NotDefined, "Не определено" },
        { (int)Types.DatesCorrection, "Перенос сроков" },
        { (int)Types.DeleteChildWork, "Удаление входящей работы" },
        { (int)Types.ChangeResource, "Изменение исполнителя" },
        { (int)Types.AddChildWork, "Добавление" },
        { (int)Types.MoveStartWorkDate, "Перенос даты начала работы" },
        { (int)Types.ChangeWorkName, "Изменение наименования" }
        };
        */
        public enum Types
        {
            NotDefined,
            DatesCorrection,
            DeleteChildWork,
            ChangeResource,
            AddChildWork,
            MoveStartWorkDate,
            ChangeWorkName
        }

        public static readonly string[] TypeNames = new string[] {
            "Не определено",
            "Перенос сроков",
            "Удаление входящей работы",
            "Изменение исполнителя",
            "Добавление",
            "Перенос даты начала работы",
            "Изменение наименования" };

        // получение описания справочника          
        static ReferenceInfo ReferenceInfo
        {
            get
            {
                return References.Connection.ReferenceCatalog.Find(EditReasonNote.ref_EditReasonNotes_Guid);
            }
        }
        // создание объекта для работы с данными
        static Reference _Reference;
        static public Reference Reference
        {
            get
            {
                if (_Reference == null)
                    _Reference = EditReasonNote.ReferenceInfo.CreateReference();
                _Reference.Refresh();
                return _Reference;
            }
        }

        // поиск типа объекта
        ClassObject classObject
        {
            get { return EditReasonNote.ReferenceInfo.Classes.Find(EditReasonNote.ERN_type_ChangeReason_Guid); }
        }

        // Справочник "Журнал редактирования планов"
        public static readonly Guid ref_EditReasonNotes_Guid = new Guid("be37cb0f-4c5c-4783-908a-fe3105fff643"); // Справочник "Журнал редактирования планов"

        public static readonly Guid ERN_type_ChangeReason_Guid = new Guid("b6591f71-9e79-4c86-92c5-7a978d8309f1"); // Тип "Причина редактирования"
        static readonly Guid ERN_param_EditType_Guid = new Guid("37e0d2d3-bbcd-43db-9f52-cd30a66b526d"); // Параметр "Тип редактирования"
        static readonly Guid ERN_param_Reason_Guid = new Guid("93e405ca-b265-4671-8229-9f5d6a67e786"); // Параметр "Причина редактирования"
        static readonly Guid ERN_param_NeedActionManager_Guid = new Guid("227f4813-7447-493a-aa4d-d2a6011823fe"); // Параметр "Требуется реакция РП" Параметр "Да/Нет"	Field_6006
        public static readonly Guid ERN_link_ToPrjctMgmntN1_Guid = new Guid("79b01004-3c10-465a-a6fb-fe2aa95ae5b8"); // Связь n:1 "Журнал редактирования планов - Управление проектами"
    }
}
