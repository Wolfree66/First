using System;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;

namespace UserRequests
{
    public class Macro : MacroProvider
    {
        public Macro(MacroContext context)
            : base(context)
        {
        }

        public override void Run()
        {
            //Test();
        }
        public void Test()
        {
            // получение описания справочника          
            ReferenceInfo referenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(new Guid("245ae93e-79cf-4a6d-a3fa-f802635e6cd1"));
            // создание объекта для работы с данными
            Reference userRequestsReference = referenceInfo.CreateReference();
            ReferenceObject ro = userRequestsReference.Find(209);
            UserRequest ur = new UserRequest(ro);
            string cat = ur.Category;
            int coef = GetCoefCategory(cat);
            this.Message(cat, coef.ToString(), null);
        }

        #region События
        public void СобытиеCохранения()
        {
            ПересчитатьПриоритет();
        }
        #endregion

        public void ПересчитатьПриоритет()
        {
            ReferenceObject referenceObject = this.Context.ReferenceObject;
            UserRequest userRequest = new UserRequest(referenceObject);
            RecalcPriority(userRequest);
        }

        private void RecalcPriority(UserRequest userRequest)
        {
            int newPriority = 0;
            if (userRequest.Category == Category.Consultation)
            { newPriority = 10; }
            else
            {
                int coefModule = userRequest.Module.Coefficient;
                //int coefModulePart = 0;
                if (userRequest.ModulePart != null) coefModule = userRequest.ModulePart.Coefficient;
                int coefCategory = GetCoefCategory(userRequest.Category);
                newPriority = userRequest.Urgency * (coefModule) + userRequest.Impact + coefCategory;
                userRequest.Priority = newPriority;
            }
            //this.Message("coefCategory", coefCategory.ToString(), null);
            userRequest.Save();
        }

        private int GetCoefCategory(string category)
        {
            int coef = 1;
            switch (category)
            {
                case Category.Incident:
                    coef = 3;
                    break;
                case Category.Problem:
                    coef = 5;
                    break;
                case Category.Consultation:
                    coef = 2;
                    break;
                default:
                    break;
            }
            return coef;
        }
    }

    struct Category
    {
        public const string Incident = "Инцидент";
        public const string Proposal = "Предложение";
        public const string Problem = "Проблема";
        public const string Consultation = "Консультация";
    }

        class UserRequest
    {
        #region Конструкторы
        public UserRequest(ReferenceObject referenceObject)
        {
            this.ReferenceObject = referenceObject;
        }
        #endregion

        #region Свойства
        public ReferenceObject ReferenceObject { get; set; }

        /// <summary>
        /// Модуль
        /// </summary>
        public Module Module
        {
            get
            {
                Module result = null;
                ReferenceObject roModule = this.ReferenceObject.GetObject(link_Module_Guid);
                if (roModule != null)
                    result = new Module(roModule);
                return result;
            }
        }
        /// <summary>
        /// Модуль
        /// </summary>
        public ModulePart ModulePart
        {
            get
            {
                ModulePart result = null;
                ReferenceObject roModule = this.ReferenceObject.GetObject(link_ModulePart_Guid);
                if (roModule != null)
                    result = new ModulePart(roModule);
                return result;
            }
        }
        int? _Priority;
        /// <summary>
        /// Приоритет
        /// </summary>
        public int Priority { get
            {
                if (_Priority == null) _Priority = this.ReferenceObject[param_Priority_Guid].GetInt16();
                return (int)_Priority;
            }
            set {
                if (_Priority != value)
                    _Priority = value;
            } }
        int? _Urgency;
        /// <summary>
        /// Срочность
        /// </summary>
        public int Urgency {
            get
            {
                if (_Urgency == null) _Urgency = this.ReferenceObject[param_Urgency_Guid].GetInt16();
                return (int)_Urgency;
            }
            set
            {
                if (_Urgency != value)
                    _Urgency = value;
            }
        }
        int? _Impact;
        /// <summary>
        /// Влияние
        /// </summary>
        public int Impact {
            get
            {
                if (_Impact == null) _Impact = this.ReferenceObject[param_Impact_Guid].GetInt16();
                return (int)_Impact;
            }
            set
            {
                if (_Impact != value)
                    _Impact = value;
            }
        }

        string _Category;
        /// <summary>
        /// Категория
        /// </summary>
        public string Category
        {
            get
            {
                if (_Category == null) _Category = this.ReferenceObject[param_Category_Guid].GetString();
                return _Category;
            }
            set
            {
                if (_Category != value)
                    _Category = value;
            }
        }

        #endregion
        #region Методы
        public void Save()
        {
            this.ReferenceObject.BeginChanges();
            this.ReferenceObject[param_Priority_Guid].Value = Priority;
            this.ReferenceObject[param_Urgency_Guid].Value = Urgency;
            this.ReferenceObject[param_Impact_Guid].Value = Impact;
            this.ReferenceObject.EndChanges();
        }
        #endregion


        static readonly Guid param_Impact_Guid = new Guid("8c138aa2-6d28-41d4-aafe-e5adbbd0ec4a");
        static readonly Guid param_Urgency_Guid = new Guid("192a233e-9ae2-45c2-bb27-2c22cb08deea");
        static readonly Guid param_Priority_Guid = new Guid("4aac1a58-b4d0-430c-bb6c-1743597130a6");
        static readonly Guid param_Category_Guid = new Guid("194d89e8-f630-4aa9-a786-92abbe3ac197");
        static readonly Guid link_Module_Guid = new Guid("522d8439-7b1a-4ae5-86fe-89766fd1b74e");
        static readonly Guid link_ModulePart_Guid = new Guid("10047835-377a-480a-b1cb-d2ae7d0ad74b");
        
    }

    class Module
    {
        #region Конструкторы
        public Module(ReferenceObject referenceObject)
        {
            this.ReferenceObject = referenceObject;
        }
        #endregion
        #region Свойства
        public ReferenceObject ReferenceObject { get; set; }

        string _Name;
        public string Name
        {
            get
            {
                if (_Name == null) _Name = this.ReferenceObject[param_Name_Guid].GetString();
                return _Name;
            }
            set
            {
                if (_Name != value)
                    _Name = value;
            }
        }
        int? _Coefficient;
        public int Coefficient {
            get
            {
                if (_Coefficient == null) _Coefficient = this.ReferenceObject[param_Coefficient_Guid].GetInt16();
                return (int)_Coefficient;
            }
            set
            {
                if (_Coefficient != value)
                    _Coefficient = value;
            }
        }
        #endregion

        static readonly Guid param_Name_Guid = new Guid("6dbf1c2a-2432-4e64-8edd-ac1d989d9677");
        static readonly Guid param_Coefficient_Guid = new Guid("1c96c1eb-1e8f-4383-b8cf-0ee5242c17a5");
    }

    class ModulePart
    {
        #region Конструкторы
        public ModulePart(ReferenceObject referenceObject)
        {
            this.ReferenceObject = referenceObject;
        }
        #endregion
        #region Свойства
        public ReferenceObject ReferenceObject { get; set; }

        string _Name;
        public string Name
        {
            get
            {
                if (_Name == null) _Name = this.ReferenceObject[param_Name_Guid].GetString();
                return _Name;
            }
            set
            {
                if (_Name != value)
                    _Name = value;
            }
        }
        int? _Coefficient;
        public int Coefficient
        {
            get
            {
                if (_Coefficient == null) _Coefficient = this.ReferenceObject[param_Coefficient_Guid].GetInt16();
                return (int)_Coefficient;
            }
            set
            {
                if (_Coefficient != value)
                    _Coefficient = value;
            }
        }
        #endregion

        static readonly Guid param_Name_Guid = new Guid("98e5c70f-df52-4f14-9212-08acf402ad92");
        static readonly Guid param_Coefficient_Guid = new Guid("2cf8bf9d-1ea7-4a9b-883d-33ee5e8d6204");
    }
}
