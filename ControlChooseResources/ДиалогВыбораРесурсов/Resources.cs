using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;

namespace ControlChooseResources
{



   public class Resource
    {
        //Справочник - "Ресурсы"
        public static readonly Guid ref_Resources_Guid = new Guid("fe80ab68-01e1-4a95-96cf-602ec877ff19");          //Guid справочника - "Ресурсы"
        public static readonly Guid type_HumanResources_Guid = new Guid("91fd9349-7045-4f26-9239-bc46be060cc0");    //Guid типа "Нерасходуемые ресурсы"
        public static readonly Guid link_Human_Resources_Guid = new Guid("b82a5812-2a71-4745-b642-72fde591be7e");   // Связь - Человек
        public static readonly Guid link_param_IncludeDate_Guid = new Guid("9cf978c2-5c86-4e23-8937-f083b3ec807a"); //Guid параметра связи "Дата включения"
        public static readonly Guid link_param_ExcludeDate_Guid = new Guid("d1e62943-073b-4725-b5f6-cc74f83eec52"); //Guid параметра связи "Дата исключения"
        public static readonly Guid item_DinamikaOrgStruct_Guid = new Guid("1ae979bd-21db-43a2-94fc-5825dee477b6"); //Guid объекта "Динамика организационная структура"

        /// <summary>
        /// Создаёт новый экземпляр ресурса
        /// </summary>
        /// <param name="current"></param>
        /// <param name="parent"></param>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        public Resource(ReferenceObject current, Resource parent, DateTime startDate, DateTime endDate)
        {
            StartUsing = startDate;
            EndUsing = endDate;
            //Children = GetActualChildren();
            Parent = parent;
            if (current != null)
            {
                Current = current;
                
            }
            else
            {
                Current = DinamikaOrgStruct;
                //Name = "";
            }
            Name = Current/*.GetObject(link_Human_Resources_Guid)*/.ToString();
            LastPlaceForWorkers = new List<Resource>();
        }

        public string Name { get; set; }

        public List<Resource> LastPlaceForWorkers { get; set; }

        /// <summary>
        /// Находит всех детей по датам включения и исключения
        /// </summary>
        /// <returns>Список существующих детей</returns>
        private List<Resource> GetAllChildren()
        {
            List<Resource> result = new List<Resource>();
            if (Current == null || !Current.HasChildren) return result;
            List<ComplexHierarchyLink> links = null;

            if (Current.Class == ClassHumanResource || Current.Class.Base == ClassHumanResource)
            {
                links = Current.Children.GetHierarchyLinks();
            }
            else return null;

            if (links == null || links.Count == 0) return result;
            ComplexHierarchyLink linkResGroup = links[0];
            //MessageBox.Show(resource.ToString() + "\n" + links.Count.ToString());
            foreach (var item in links)
            {
                    Resource child = new Resource(item.ChildObject, this, item[link_param_IncludeDate_Guid].GetDateTime(), item[link_param_ExcludeDate_Guid].GetDateTime());
                    result.Add(child);
                if (IsLastPlaceOfWork(child)) LastPlaceForWorkers.Add(child);
            }
            result.Sort(delegate (Resource x, Resource y)
            {
                if (x.Name == null && y.Name == null) return 0;
                else if (x.Name == null) return -1;
                else if (y.Name == null) return 1;
                else return x.Name.CompareTo(y.Name);
            });
            //if(linkResGroup == null) return links[0].ParentObject;
            return result;
           // throw new NotImplementedException("Ошибка нахождения дочерних элементов");
        }

        private bool IsLastPlaceOfWork(Resource childResource)
        {
            //if (child.Name.Contains("Денисов"))
           //     new WarningException("");
            DateTime latestDate = childResource.EndUsing;
            //DateTime tempDate = latestDate;
            if (latestDate == DateTime.MinValue || childResource.Children.Count != 0) return false;
            
            foreach (var item in childResource.Current.Parents.GetHierarchyLinks())
            {
                DateTime excludeDate = item[link_param_ExcludeDate_Guid].GetDateTime();
                if (excludeDate == DateTime.MinValue || excludeDate > latestDate) return false; 
            }
            return true;
        }

        public override string ToString()
        {
            return Name;
        }

       /* public bool IsActual { get
            {
                if (EndUsing == null || EndUsing == DateTime.MinValue) return true;
                if (EndUsing < DateTime.Today) return false;
                return true;
            } }
*/
        public ReferenceObject Current { get; set; }

        public Resource Parent { get; set; }

        List<Resource> _Children;
        public List<Resource> Children {
            get
            {
                List<Resource> temp = new List<Resource>();
                if (_Children == null)
                {
                    _Children = GetAllChildren();
                }
                return _Children;
            }
            set { } }

        public DateTime StartUsing { get; set; }

        public DateTime EndUsing { get; set; }

        private Reference _ResourcesReference;
        private static ReferenceInfo _ResourcesReferenceInfo;
        private static ClassObject _humanResource;
        private static ReferenceObject _DinamikaOrgStruct;
        public static ServerConnection Connection
        { get { return ServerGateway.Connection; } set {/*Connection = value; */} }

        private static Reference GetReference(ref Reference reference, ReferenceInfo referenceInfo)
        {
            if (reference == null)
                reference = referenceInfo.CreateReference();


            return reference;
        }

        /// <summary>
        /// Справочник "Ресурсы"
        /// </summary>
        private Reference ResourcesReference
        {
            get
            {
                if (_ResourcesReference == null)
                {
                    _ResourcesReference = GetReference(ref _ResourcesReference, ResourcesReferenceInfo);
                    _ResourcesReference.LoadSettings.LoadDeleted = false;
                    _ResourcesReference.LoadSettings.AddParameters(new Guid [] { Resource.link_param_IncludeDate_Guid, Resource.link_param_ExcludeDate_Guid});
                    _ResourcesReference.Objects.Load();
                }

                return _ResourcesReference;
            }
        }

        private ReferenceInfo GetReferenceInfo(ref ReferenceInfo referenceInfo, Guid referenceGuid)
        {
            if (referenceInfo == null)
                referenceInfo = Connection.ReferenceCatalog.Find(referenceGuid);

            return referenceInfo;
        }

        private ReferenceInfo ResourcesReferenceInfo
        {
            get { return GetReferenceInfo(ref _ResourcesReferenceInfo, ref_Resources_Guid); }
        }

        /// <summary>
        /// Объект - Динамика организационная структура
        /// </summary>
        public ReferenceObject DinamikaOrgStruct
        {
            get
            {
                if (_DinamikaOrgStruct == null)
                    _DinamikaOrgStruct = ResourcesReference.Find(item_DinamikaOrgStruct_Guid);

                return _DinamikaOrgStruct;
            }
        }

        /// <summary>
        /// класс нерасходуемый ресурс
        /// </summary>
        private ClassObject ClassHumanResource
        {
            get
            {
                if (_humanResource == null)
                    _humanResource = ResourcesReference.Classes.Find(type_HumanResources_Guid);

                return _humanResource;
            }
        }
    }
}
