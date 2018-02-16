using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model.References;

namespace WpfApp_DialogueAddSignatories.Model
{
    /// <summary>
    /// Для работы с используемыми ресурсами в УП
    /// ver 1.0
    /// </summary>
    public class UsedResource
    {
        public UsedResource(ReferenceObject usedResource_ro)
        {
            this.ReferenceObject = usedResource_ro;
            this.StartDate = usedResource_ro[param_UsResStartDate_Guid].GetDateTime();
            this.EndDate = usedResource_ro[param_UsResEndDate_Guid].GetDateTime();
            this.Resource = usedResource_ro.GetObject(link_Resource_Guid);
            this.LabourHours = usedResource_ro[param_NumHours_Guid].GetDouble();
        }

        public UsedResource(ReferenceObject resource, double labourHours, DateTime startDate, DateTime endDate)
        {
            this.ReferenceObject = References.UsedResources.CreateReferenceObject(References.Class_NonConsumableResources);
            //this.ReferenceObject.ApplyChanges();
            this.Resource = resource;
            this.StartDate = startDate;
            this.EndDate = endDate;
            this.LabourHours = labourHours;
            SaveToDB();
        }

        //private void LinkToWork(ProjectManagementWork work)
        //{
        //    this.ReferenceObject.BeginChanges();
        //    this.ReferenceObject.SetLinkedObject(link_Resource_Guid, work.ReferenceObject);
        //    this.ReferenceObject.EndChanges();
        //}

        void SaveToDB()
        {
            //this.ReferenceObject.BeginChanges();
            this.ReferenceObject[param_NumIsFixed_Guid].Value = true;
            this.ReferenceObject[param_UsResStartDate_Guid].Value = this.StartDate;
            this.ReferenceObject[param_UsResEndDate_Guid].Value = this.EndDate;
            this.ReferenceObject.SetLinkedObject(link_Resource_Guid, this.Resource);
            this.ReferenceObject[param_NumHours_Guid].Value = this.LabourHours;
            //MessageBox.Show("start SaveToDB");
            bool savig = this.ReferenceObject.EndChanges();
            //MessageBox.Show("SaveToDB" + savig.ToString());
        }

        internal void Delete()
        {
            this.ReferenceObject.Delete();
        }

        ProjectManagementWork _Work;
        public ProjectManagementWork Work
        {
            get
            {
                if (_Work == null)
                {
                    ReferenceObject work_ro = this.ReferenceObject.GetObject(link_UsedResources_GUID);
                    if (work_ro != null)
                        _Work = new ProjectManagementWork(work_ro);
                }
                return _Work;
            }
        }

        bool? _IsPlanned;
        public bool IsPlanned
        {
            get
            {
                if (_IsPlanned == null)
                    _IsPlanned = !this.ReferenceObject[param_IsPlanned_Guid].GetBoolean();
                return (bool)_IsPlanned;
            }
        }
        public ReferenceObject ReferenceObject
        {
            get; private set;
        }

        public DateTime StartDate
        { get; private set; }

        public DateTime EndDate
        { get; private set; }

        public ReferenceObject Resource
        { get; private set; }

        public double LabourHours
        { get; private set; }

        private static readonly Guid param_UsResStartDate_Guid = new Guid("57695721-084c-48bf-8c39-667d27ee1aaf");     //Guid параметра "Начало"
        private static readonly Guid param_UsResEndDate_Guid = new Guid("1680be8c-8527-4243-85ab-b3ae6dc38140");       //Guid параметра "Окончание"
        private static readonly Guid param_NumHours_Guid = new Guid("28076069-6b20-4b66-97ab-a03e2b14f655");           //Guid параметра "Количество"
        private static readonly Guid param_NumIsFixed_Guid = new Guid("98b3cfd1-914e-402a-a05c-571d15ceab6a");         //Guid параметра "Фиксировать количество"
        private static readonly Guid param_PlanningSpace_Guid = new Guid("01b7d086-1483-42ee-bc14-588dbabd7d50");   //Guid параметра "Пространство планирования"
        private static readonly Guid param_IsPlanned_Guid = new Guid("4227e515-66a4-43ae-9418-346854748986");         //Guid параметра "Фактическое значение
        private static readonly Guid link_ResourceGroup_Guid = new Guid("35edec01-6119-44c5-a062-1859eeac38ab");       //Guid связи N:1 "Группа"
        private static readonly Guid link_Resource_Guid = new Guid("7f882c52-bfae-4a93-a7a7-9f548215898f");            //Guid связи n:1 "Ресурс"
        private static readonly Guid link_UsedResources_GUID = new Guid("1a22ee46-5438-4caa-8b75-8a8a37b74b7e");       //Guid связи 1:n "Ресурсы" (используемы ресурсы в работе)
    }
}
