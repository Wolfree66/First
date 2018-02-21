//#define server
//#define test
namespace ResolvingConflicts
{
    using System;
    using System.Collections.Generic;
    using TFlex.DOCs.Model.Macros;
    using TFlex.DOCs.Model.References;
    //using TFlex.DOCs.Model.References.Workflow;
    using TFlex.DOCs.Model.Processes.Events.Contexts;
    using TFlex.DOCs.Model.Processes.Events.Contexts.Data;
    using TFlex.DOCs.Model.Mail;
    using TFlex.DOCs.Model.References.Processes;
    using TFlex.DOCs.Model.References.Users;
    using TFlex.DOCs.Model;
    using System.Text;
    using System.Linq;
    using TFlex.DOCs.Model.References.Files;
    using System.IO;
    using TFlex.DOCs.Model.Stages;
#if !server

    using System.Windows.Forms;
    using TFlex.DOCs.Model.References.Procedures;
    using TFlex.DOCs.UI.Objects.Managers;
#endif

    public class Macro : MacroProvider
    {
        public Macro(MacroContext context)
            : base(context)
        {
        }


        //private Guid declarantManagers_link_Guid = new Guid("a8ced5e5-c8d4-4c70-a54c-fecbf4a959e1"); //связь Руководитель заявителя (группы и пользователи)
        private Guid executorManagers_link_Guid = new Guid("64ff13ea-2dac-4cc4-8def-e79340be0642"); //связь Начальник отдела (с ИО) исполнителя
        private Guid executorBoss_link_Guid = new Guid("ecb3497e-d186-4c7d-8ba1-a74f64734b2d"); //связь Руководитель подразделения исполнителя
        private Guid executorViceBoss_link_Guid = new Guid("0df30206-f00a-43fc-bd2d-5d21ce834dc2"); //связь Заместитель Руководителя подразделения исполнителя
        private Guid executor_param_Guid = new Guid("755741dc-89ed-4e95-bc06-7dbe55032814"); //связь Противоречие - Исполнитель
        private Guid guid_nach_isp = new Guid("64ff13ea-2dac-4cc4-8def-e79340be0642"); // Гуид cвязи Начальник отдела (с ИО)
        private Guid guid_isp_name = new Guid("755741dc-89ed-4e95-bc06-7dbe55032814"); // Исполнитель-Противоречие
        private Guid guid_ruk_podr_name = new Guid("ecb3497e-d186-4c7d-8ba1-a74f64734b2d"); // Гуид связи Руководитель подразделения исполнителя
        private Guid guid_zam_nach_isp = new Guid("0df30206-f00a-43fc-bd2d-5d21ce834dc2"); // Гуид cвязи Заместитель Руководителя подразделения исполнителя
        private Guid guid_podr_isp = new Guid("21ee9f6a-3b1c-4986-8475-0bc93b699c4e"); // Гуид подразделения исполнителя
        private Guid guid_z_name = new Guid("0db861e1-e963-44b2-a727-e12149f42db1"); // Заявитель
        private Guid guid_ruk_otd_z_name = new Guid("a8ced5e5-c8d4-4c70-a54c-fecbf4a959e1"); // Гуид связи Руководитель заявителя (группы и пользователи)
        private Guid guid_name = new Guid("beb2d7d1-07ef-40aa-b45b-23ef3d72e5aa"); // Гуид наименования ГиП
        private Guid guid_GIP = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea"); // Гуид ГиП
        private Guid guid_GIP_komm = new Guid("c96c2bbf-8563-47c0-8842-a1e2cc42092a"); // Гуид ГиП Группа пользователей-Комментарий
                                                                                       //private Guid guid_ruk_otd_name = new Guid ("fdb41549-2adb-40b0-8bb5-d8be4a6a8cd2"); // ГиП - Руководитель группы
        private Guid guid_zam_ruk_otd_name = new Guid("a3529004-af33-4582-916c-93c9ac6fb481"); // ГиП - заместитель начальника подразделения (сделать связь, диалог и скопировать гуид на боевом сервере)
        private Guid performPosition_Link_Guid = new Guid("59921b0b-aec2-4282-a1bd-e9ee9e57a8ab"); // связь Должность Исполнитель

        //Guid справочника - "Процедуры"
        private readonly Guid referenceProceduresGUID = new Guid("61d922d0-4d60-49f1-a6a0-9b6d22d3d8b3");
        //Guid связи n:n "Запущенные объекты"
        private readonly Guid workingObjectsGUID = new Guid("cf72b950-52d9-4089-a07b-efad09ed613c");

        public override void Run()
        {

        }

        public void Test()
        {
            // получение описания справочника          
            ReferenceInfo referenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(new Guid("e8aebfce-ba8d-4b98-98af-aea92320447b"));
            // создание объекта для работы с данными
            Reference reference = referenceInfo.CreateReference();
            ReferenceObject referenceObject = reference.Find(378);
            Conflict conflict = new ResolvingConflicts.Conflict(referenceObject);
            SendMessageAboutNewConflict(conflict);
            return;
            StringBuilder strB = new StringBuilder();
            foreach (var item in conflict.DeclarantManagers)
            {
                strB.AppendLine(item.ToString());
            }
            strB.AppendLine(conflict.Performer.ToString());
        }



        #region Точки входа

        #region События справочника
#if !server
        public void Событие_НажатиеНаКнопку_ОтправитьОтветЗаявителю()
        {
            ReferenceObject referenceObject = Context.ReferenceObject;
            Conflict conflict = new Conflict(referenceObject);
            ChangeStage(conflict, StageName.AnswerReceived);
            referenceObject.ApplyChanges();
            SendMessageAboutPerformerAnswer(conflict);
            ЗакрытьДиалогСвойств();
            //MessageBox.Show("Ответ отослан заявителю");
        }

        public void ЗакрытьДиалогСвойств()
        {
            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true);
        }

#endif
        public void Событие_ЗавершениеСохраненияОбъекта()
        {
            ReferenceObject referenceObject = Context.ReferenceObject;
            Conflict conflict = new Conflict(referenceObject);
            ChangeStage(conflict, StageName.NewConflict);
#if !server
            ЗапуститьПроцесс();
#endif
        }

        #endregion

        public void SendMessagesAboutNewConflict()
        {
            EventContext КонтекстСобытияБП = Context as EventContext;//контекст событий по БП
            if (КонтекстСобытияБП != null)
            {
                //КонтекстСобытияБП = Context as EventContext;//контекст событий по БП
                EventContextData ДанныеЭтапа = КонтекстСобытияБП.Data;
                ProcessReferenceObject proc = ДанныеЭтапа.Process;
                List<ReferenceObject> listObj = proc.GetObjects(workingObjectsGUID);
                foreach (ReferenceObject item in listObj)
                {
                    Conflict conflict = new Conflict(item);
                    SendMessageAboutNewConflict(conflict);

                }
            }
        }



        public void SendMessagesDeclarantManagers()
        {
            SendMessagesAboutNewConflict();
            //EventContext КонтекстСобытияБП = Context as EventContext;//контекст событий по БП
            //if (КонтекстСобытияБП != null)
            //{
            //    //КонтекстСобытияБП = Context as EventContext;//контекст событий по БП
            //    EventContextData ДанныеЭтапа = КонтекстСобытияБП.Data;
            //    ProcessReferenceObject proc = ДанныеЭтапа.Process;
            //    List<ReferenceObject> listObj = proc.GetObjects(workingObjectsGUID);
            //    foreach (ReferenceObject item in listObj)
            //    {

            //        MailMessage message = new MailMessage(DOCsAccount.Instance); //Создаем новое сообщение
            //        message.Subject = "Создано новое противоречие вашим подчинённым.";
            //        message.Body = "В модуле \"Решение противоречий\" ваш подчинённый - "
            //            + item.SystemFields.Author.ToString()
            //            + " создал новое противоречие. Ознакомьтесь с противоречием, перейдя по вложенной вкладке.";
            //        List<ReferenceObject> users = item.GetObjects(declarantManagers_link_Guid);
            //        if (users == null) return;
            //        foreach (ReferenceObject user in users)
            //        {
            //            //Добавляем адресатов сообщения
            //            User mailuser = user as User;
            //            message.To.Add(new MailUser(mailuser));
            //        }
            //        message.Attachments.Add(new ObjectAttachment(item));//Прикрепляем к сообщению объект справочника
            //        message.Send(); //Отправляем сообщение
            //    }
            //}
        }

        public void SendMessagesDeclarantManagersAboutResponse()
        {
            EventContext КонтекстСобытияБП = Context as EventContext;//контекст событий по БП
            if (КонтекстСобытияБП != null)
            {
                //КонтекстСобытияБП = Context as EventContext;//контекст событий по БП
                EventContextData ДанныеЭтапа = КонтекстСобытияБП.Data;
                ProcessReferenceObject proc = ДанныеЭтапа.Process;
                List<ReferenceObject> listObj = proc.GetObjects(workingObjectsGUID);
                foreach (ReferenceObject item in listObj)
                {
                    Conflict conflict = new Conflict(item);
                    SendMessageAboutResponse(conflict);
                }
            }
        }

        public void SendMessagesExecutorManagers()
        {
            EventContext КонтекстСобытияБП = Context as EventContext;//контекст событий по БП
            if (КонтекстСобытияБП != null)
            {
                //КонтекстСобытияБП = Context as EventContext;//контекст событий по БП
                EventContextData ДанныеЭтапа = КонтекстСобытияБП.Data;
                ProcessReferenceObject proc = ДанныеЭтапа.Process;
                List<ReferenceObject> listObj = proc.GetObjects(workingObjectsGUID);
                foreach (ReferenceObject item in listObj)
                {

                    MailMessage message = new MailMessage(DOCsAccount.Instance); //Создаем новое сообщение
                    message.Subject = "Создано новое противоречие в котором указан исполнителем ваш подчинённый.";
                    message.Body = "В модуле \"Решение противоречий\" ваш подчинённый - "
                        + item.GetObject(executor_param_Guid).ToString()
                        + " указан как исполнитель. Ознакомьтесь с противоречием, перейдя по вложенной вкладке и примите меры по ее устранению.";
                    List<ReferenceObject> users = item.GetObjects(executorManagers_link_Guid);
                    if (users == null) return;
                    foreach (ReferenceObject user in users)
                    {
                        //Добавляем адресатов сообщения
                        User mailuser = user as User;
                        message.To.Add(new MailUser(mailuser));
                    }
                    message.Attachments.Add(new ObjectAttachment(item));//Прикрепляем к сообщению объект справочника
                    message.Send(); //Отправляем сообщение
                }
            }
        }

        public void SendMessagesExecutorBosses()
        {
            EventContext КонтекстСобытияБП = Context as EventContext;//контекст событий по БП
            if (КонтекстСобытияБП != null)
            {
                //КонтекстСобытияБП = Context as EventContext;//контекст событий по БП
                EventContextData ДанныеЭтапа = КонтекстСобытияБП.Data;
                ProcessReferenceObject proc = ДанныеЭтапа.Process;
                List<ReferenceObject> listObj = proc.GetObjects(workingObjectsGUID);
                foreach (ReferenceObject item in listObj)
                {

                    MailMessage message = new MailMessage(DOCsAccount.Instance); //Создаем новое сообщение
                    message.Subject = "Нет ответа на противоречие в котором указан исполнителем ваш подчинённый.";
                    message.Body = "В модуле Решение противоречий 3 часа назад создано новое противоречие, которое на данный момент находится без ответа и в котором Ваш подчиненный  - "
                        + item.GetObject(executor_param_Guid).ToString()
                        + " указан как исполнитель. Ознакомьтесь с противоречием, перейдя по вложенной вкладке и примите меры по ее устранению.";
                    List<ReferenceObject> users = item.GetObjects(executorBoss_link_Guid);
                    if (users == null) return;
                    foreach (ReferenceObject user in users)
                    {
                        //Добавляем адресатов сообщения
                        User mailuser = user as User;
                        message.To.Add(new MailUser(mailuser));
                    }
                    users = item.GetObjects(executorViceBoss_link_Guid);
                    if (users == null) return;
                    foreach (ReferenceObject user in users)
                    {
                        //Добавляем адресатов сообщения
                        User mailuser = user as User;
                        message.To.Add(new MailUser(mailuser));
                    }
                    message.Attachments.Add(new ObjectAttachment(item));//Прикрепляем к сообщению объект справочника
                    message.Send(); //Отправляем сообщение
                }
            }
        }
        #endregion





#if !server
        public void ЗапуститьПроцесс()
        {
            ReferenceObject противоречие = Context.ReferenceObject;
            if (противоречие.SystemFields.Stage.ToString() == "Новое противоречие")
            {
                if (MessageBox.Show(("Запустить процесс Решения противоречий?"),
                                   "Запустить процесс Решения противоречий?", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.Cancel)
                { return; }
                else
                {
                    RunProcess();
                    //MessageBox.Show("процесс запущен.");
                }
            }
        }

        public void RunProcess()
        {
            ServerConnection servConnect = (Context as MacroContext).Connection;
            // получение описания справочника "Процедуры"
            ReferenceInfo referenceInfo = servConnect.ReferenceCatalog.Find(referenceProceduresGUID);
            // создание объекта для работы с данными
            Reference reference = referenceInfo.CreateReference();
            ReferenceObject procedureRef = reference.Find(new Guid("a8ce3994-a144-4227-9f4b-3e4ea4cef704"));//гуид процедуры "Схема Решение противоречий"
            ProcedureReferenceObject proc = procedureRef as ProcedureReferenceObject;

            ReferenceObject current = Context.ReferenceObject;

            CreateWorkflowContext workflowContext = new CreateWorkflowContext(proc, new List<ReferenceObject> { current });

            workflowContext.ProcessName = "Решение противоречий";//присваиваем название процессу = параметр "Номер поручения"

            //ProcessReferenceObject newProcess = proc.GenerateProcess(workflowContext);
            bool newProcess = proc.CreateProcess(workflowContext);

            if (newProcess) MessageBox.Show("Процесс - запущен.");
            else MessageBox.Show("Невозможно запустить Процесс, возможно некорректно задан Исполнитель.");
        }


        public void Определение_параметров_исполнителя()
        {
            ReferenceObject противоречие = Context.ReferenceObject;
            Conflict conflict = new ResolvingConflicts.Conflict(противоречие);

            UserReferenceObject исполнитель = conflict.Performer; // Исполнитель-Противоречие

            var должность = ВыполнитьМакрос("6a4385ff-475a-4afb-bb85-965252abebdc", "ОбъектДолжность", исполнитель);

            ServerConnection servConnect = (Context as MacroContext).Connection;
            // получение описания справочника
            //ReferenceInfo referenceInfo = servConnect.ReferenceCatalog.Find(refUsers);
            // создание объекта для работы с данными
            UserReference reference = new UserReference(servConnect);


            //MessageBox.Show(списокОтделовЗаявителя.Count.ToString());
            //bool send = false;
            ReferenceObject departmentObj = null;
            ReferenceObject groupObj = должность.ParentObject;
            List<ReferenceObject> departmentManagers = new List<ReferenceObject>();
            List<ReferenceObject> groupManagers = new List<ReferenceObject>();
            List<ReferenceObject> departmentViceManagers = new List<ReferenceObject>();
            List<ReferenceObject> groupViceManagers = new List<ReferenceObject>();

            var списокОтделов = ВыполнитьМакрос("60d73bc1-bb67-4730-8877-eb6e0f93837a", "GetTheStackPath", groupObj);

            if (списокОтделов.Count > 1)
            {
                for (int i = 0; i <= списокОтделов.Count - 2; i++)
                {
                    Guid objectGuid2 = new Guid(списокОтделов[i].Параметр["GUID"].ToString());
                    ReferenceObject departmentItem = reference.Find(objectGuid2);
                    //MessageBox.Show(authorDepartmentItem.ToString());
                    if ((GetTheBosses(departmentItem).Count > 0 || GetTheViceBosses(departmentItem).Count > 0) && i < списокОтделов.Count - 2)
                    {
                        //MessageBox.Show("add manager");
                        groupManagers = GetTheBosses(departmentItem);
                        groupViceManagers = GetTheViceBosses(departmentItem);
                        i = списокОтделов.Count - 2;
                        groupObj = departmentItem;
                    }
                    if (i == списокОтделов.Count - 2)
                    {
                        objectGuid2 = new Guid(списокОтделов[i].Параметр["GUID"].ToString());
                        departmentItem = reference.Find(objectGuid2);
                        //MessageBox.Show("add depart manager");
                        departmentManagers = GetTheBosses(departmentItem);
                        departmentViceManagers = GetTheViceBosses(departmentItem);
                        departmentObj = departmentItem;
                    }
                }
            }
            противоречие.BeginChanges();
            //связь с отделом исполнителя
            if (groupObj != null)
            {
                противоречие.ParameterValues[new Guid("21ee9f6a-3b1c-4986-8475-0bc93b699c4e")].Value = groupObj.ToString();
                if (groupManagers.Count > 0)
                {
                    foreach (ReferenceObject item in groupManagers)
                    {
                        //MessageBox.Show("add manager " + item.ToString());
                        противоречие.AddLinkedObject(guid_nach_isp, item);
                    }
                }
                if (groupViceManagers.Count > 0)
                {
                    foreach (ReferenceObject item in groupViceManagers)
                    {
                        //MessageBox.Show("add vice manager " + item.ToString());
                        противоречие.AddLinkedObject(guid_nach_isp, item);
                    }
                }
                if (departmentManagers.Count > 0)
                {
                    foreach (ReferenceObject item in departmentManagers)
                    {
                        //MessageBox.Show("add department manager " + item.ToString());
                        противоречие.AddLinkedObject(guid_ruk_podr_name, item);
                    }
                }
                if (departmentViceManagers.Count > 0)
                {
                    foreach (ReferenceObject item in departmentViceManagers)
                    {
                        //MessageBox.Show("add department vice manager " + item.ToString());
                        противоречие.AddLinkedObject(guid_zam_nach_isp, item);
                    }
                }
            }
            else
            {
                противоречие.ParameterValues[new Guid("21ee9f6a-3b1c-4986-8475-0bc93b699c4e")].Value = departmentObj.ToString();
                if (departmentManagers.Count > 0)
                {
                    foreach (ReferenceObject item in departmentManagers)
                    {
                        //MessageBox.Show("add department manager " + item.ToString());
                        противоречие.AddLinkedObject(guid_nach_isp, item);
                        противоречие.AddLinkedObject(guid_ruk_podr_name, item);
                    }
                }
                if (departmentViceManagers.Count > 0)
                {
                    foreach (ReferenceObject item in departmentViceManagers)
                    {
                        //MessageBox.Show("add department vice manager " + item.ToString());
                        противоречие.AddLinkedObject(guid_nach_isp, item);
                        противоречие.AddLinkedObject(guid_zam_nach_isp, item);
                    }
                }
            }
            //связь с подразделением исполнителя
            противоречие.ParameterValues[new Guid("13d91e17-c410-4933-86e9-333b1a0f7e1b")].Value = departmentObj.ToString();
            //связь с должностью исполнителя
            противоречие.SetLinkedObject(new Guid("{59921b0b-aec2-4282-a1bd-e9ee9e57a8ab}"), groupObj);


            противоречие.EndChanges();
        }

        public void Определение_параметров_заявителя()
        {
            User currentUser = ServerGateway.Connection.ClientView.GetUser();
            //          Объект заявительОбъект = ТекущийПользователь;
            //RefObj заявитель_RefObj = CurrentUser;
            var заявительДолжность = ВыполнитьМакрос("6a4385ff-475a-4afb-bb85-965252abebdc", "ОбъектДолжность", currentUser);
            //ReferenceObject authorPosition = заявительДолжность as ReferenceObject;
            ReferenceObject противоречие = Context.ReferenceObject;

            ServerConnection servConnect = (Context as MacroContext).Connection;

            ReferenceObject authorGroup = заявительДолжность.ParentObject;

            // установка связи 1:1 или N:1



            //MessageBox.Show(authorPosition.ToString());
            var списокОтделовЗаявителя = ВыполнитьМакрос("60d73bc1-bb67-4730-8877-eb6e0f93837a", "GetTheStackPath", authorGroup);
            //MessageBox.Show(списокОтделовЗаявителя.Count.ToString());
            //bool send = false;
            List<ReferenceObject> authorsDepartmentManagers = new List<ReferenceObject>();
            List<ReferenceObject> authorsGroupManagers = new List<ReferenceObject>();
            List<ReferenceObject> authorsDepartmentViceManagers = new List<ReferenceObject>();
            List<ReferenceObject> authorsGroupViceManagers = new List<ReferenceObject>();
            if (списокОтделовЗаявителя.Count > 1)
            {
                for (int i = 0; i <= списокОтделовЗаявителя.Count - 2; i++)
                {
                    //Guid objectGuid2 = new Guid(списокОтделовЗаявителя[i].Параметр["GUID"].ToString());
                    ReferenceObject authorDepartmentItem = (ReferenceObject)списокОтделовЗаявителя[i];
                    //MessageBox.Show(authorDepartmentItem.ToString());
                    if ((GetTheBosses(authorDepartmentItem).Count > 0 || GetTheViceBosses(authorDepartmentItem).Count > 0) && i < списокОтделовЗаявителя.Count - 2)
                    {
                        //MessageBox.Show("add manager");
                        authorsGroupManagers = GetTheBosses(authorDepartmentItem);
                        authorsGroupViceManagers = GetTheViceBosses(authorDepartmentItem);
                        i = списокОтделовЗаявителя.Count - 2;
                    }
                    if (i == списокОтделовЗаявителя.Count - 2)
                    {
                        //MessageBox.Show("add depart manager");
                        //objectGuid2 = new Guid(списокОтделовЗаявителя[i].Параметр["GUID"].ToString());
                        authorDepartmentItem = (ReferenceObject)списокОтделовЗаявителя[i];
                        //MessageBox.Show("add depart manager");
                        authorsDepartmentManagers = GetTheBosses(authorDepartmentItem);
                        authorsDepartmentViceManagers = GetTheViceBosses(authorDepartmentItem);
                        //departmentObj = authorDepartmentItem;
                    }
                }
            }
            противоречие.BeginChanges();
            //связь с подразделением Заявителя
            противоречие.ParameterValues[new Guid("f3a17147-5316-42a1-ad62-c46e22c6ad46")].Value = списокОтделовЗаявителя[0].ToString();
            //связь с должностью Заявителя
            противоречие.SetLinkedObject(new Guid("{136091fc-a834-4488-9ec4-4db4dd61ab71}"), authorGroup);
            if (authorsGroupManagers.Count > 0)
            {
                foreach (ReferenceObject item in authorsGroupManagers)
                {
                    //MessageBox.Show("add manager " + item.ToString());
                    противоречие.AddLinkedObject(guid_ruk_otd_z_name, item);
                }
            }
            if (authorsGroupViceManagers.Count > 0)
            {
                foreach (ReferenceObject item in authorsGroupViceManagers)
                {
                    //MessageBox.Show("add vice manager " + item.ToString());
                    противоречие.AddLinkedObject(guid_ruk_otd_z_name, item);
                }
            }
            if (authorsGroupViceManagers.Count <= 0 && authorsGroupManagers.Count <= 0)
            {
                foreach (ReferenceObject item in authorsDepartmentManagers)
                {
                    //MessageBox.Show("add manager " + item.ToString());
                    противоречие.AddLinkedObject(guid_ruk_otd_z_name, item);
                }
                foreach (ReferenceObject item in authorsDepartmentViceManagers)
                {
                    //MessageBox.Show("add manager " + item.ToString());
                    противоречие.AddLinkedObject(guid_ruk_otd_z_name, item);
                }
            }
            противоречие.EndChanges();

        }

        //Возвращает список начальников и замов для группы пользователей
        public List<ReferenceObject> GetTheBosses(ReferenceObject department)
        {
            Guid boss_Link_Child_guid = new Guid("fdb41549-2adb-40b0-8bb5-d8be4a6a8cd2");
            List<ReferenceObject> bosses = new List<ReferenceObject>();
            if (department.GetObject(boss_Link_Child_guid) != null)//Если для группы задан руководитель
            {
                //MessageBox.Show(authorDepartmentItem.ToString()+ " Есть начальник");
                if (department.GetObject(boss_Link_Child_guid).Class.ToString() == "Сотрудник"
                   || department.GetObject(boss_Link_Child_guid).Class.ToString() == "Администратор")
                {
                    //add the authors boss
                    bosses.Add(department.GetObject(boss_Link_Child_guid));
                }
                if (department.GetObject(boss_Link_Child_guid).Class.ToString() == "Должность")
                {
                    foreach (ReferenceObject item in department.GetObject(boss_Link_Child_guid).Children)
                    {
                        //MessageBox.Show(item.ToString() + "; "+item.Class.ToString() + "; Children");
                        if (item.Class.ToString() == "Сотрудник" || item.Class.ToString() == "Администратор")
                        {
                            bosses.Add(item);
                        }
                    }
                }

            }

            return bosses;
        }

        //Возвращает список начальников и замов для группы пользователей
        public List<ReferenceObject> GetTheViceBosses(ReferenceObject department)
        {
            Guid viceBoss_Link_guid = new Guid("a3529004-af33-4582-916c-93c9ac6fb481");
            List<ReferenceObject> bosses = new List<ReferenceObject>();

            if (department.GetObject(viceBoss_Link_guid) != null)//заместитель Руководителя
            {
                if (department.GetObject(viceBoss_Link_guid).Class.ToString() == "Сотрудник"
                   || department.GetObject(viceBoss_Link_guid).Class.ToString() == "Администратор")
                {
                    //add the authors boss
                    bosses.Add(department.GetObject(viceBoss_Link_guid));
                    //if(authorPositionStack[i].GetObject("fdb41549-2adb-40b0-8bb5-d8be4a6a8cd2")!=null)
                }
                if (department.GetObject(viceBoss_Link_guid).Class.ToString() == "Должность")
                {
                    foreach (var item in department.GetObject(viceBoss_Link_guid).Children)
                    {
                        bosses.Add(item);
                    }
                }

            }
            return bosses;
        }

        public void ИзменениеИсполнителя()
        {
            ReferenceObject currentObject = Context.ReferenceObject;
            var uiContext = Context as UIMacroContext;


            if (uiContext.ChangedLink.ToString() == "Противоречие - Исполнитель" && currentObject.GetObject(guid_isp_name) != null)
            {
                currentObject.BeginChanges();
                //currentObject.SetLinkedObject(guid_isp_name, null); //обнуляем связь противоречие исполнитель
                currentObject.SetLinkedObject(performPosition_Link_Guid, null); //обнуляем связь должность исполнителя
                currentObject.ClearLinks(guid_zam_nach_isp); //обнуляем связь заместитель руководителя подразделения исполнителя
                currentObject.ClearLinks(guid_ruk_podr_name); //обнуляем связь заместитель руководителя подразделения исполнителя
                currentObject.EndChanges();
                //MessageBox.Show("ispolnitel changed");
                Определение_параметров_исполнителя();
                //Обновление_Удостоверения_всем();//обновляем значения в Командировочных удостоверениях по данным из заявки
                //ОбновлениеСлужебныхЗаданийВсем();//обновляем значения в Служебных Заданиях по данным из заявки
            }
            /*else
                {
                currentObject.BeginChanges();
                //currentObject.SetLinkedObject(guid_isp_name, null); //обнуляем связь противоречие исполнитель
                currentObject.SetLinkedObject(performPosition_Link_Guid, null); //обнуляем связь должность исполнителя
                currentObject.ClearLinks(guid_zam_nach_isp); //обнуляем связь заместитель руководителя подразделения исполнителя
                currentObject.ClearLinks(guid_ruk_podr_name); //обнуляем связь заместитель руководителя подразделения исполнителя
                currentObject.EndChanges();
            }*/
        }
#endif
        private void SendMessageAboutNewConflict(Conflict conflict)
        {
            List<UserReferenceObject> mailRecipients = new List<UserReferenceObject> { conflict.Performer };

            var declarantManagers = conflict.DeclarantManagers;
            if (declarantManagers.Count() == 0) throw new ArgumentNullException("Не найдены руководители исполнителя");
            List<UserReferenceObject> mailCopyRecipients = new List<UserReferenceObject>();
            foreach (var declarantManager in declarantManagers)
            {
                mailCopyRecipients.Add(declarantManager);
            }
            string subject = "Создано новое противоречие.";
            string bodyText = CreateHTMLBodyText(conflict);

            SendHTMLMessage(subject, bodyText, mailRecipients, mailCopyRecipients, new List<Conflict> { conflict });
        }

        private void SendMessageAboutResponse(Conflict conflict)
        {
            List<UserReferenceObject> mailCopyRecipients = new List<UserReferenceObject>();

            var declarantManagers = conflict.DeclarantManagers;
            if (declarantManagers.Count() == 0) throw new ArgumentNullException("Не найдены руководители заявителя");
            List<UserReferenceObject> mailRecipients = new List<UserReferenceObject>();
            foreach (var declarantManager in declarantManagers)
            {
                mailCopyRecipients.Add(declarantManager);
            }
            string subject = "Ответ на противоречие получен.";
            string bodyText = "В модуле 'Решение противоречий' вашим подчиненным - "
                + conflict.Declarant.ToString() + " - было создано противоречие "
                + conflict.Category + ". На него был дан ответ:" + conflict.Answer
                + ". При необходимости закройте этот вопрос. Для этого выберите правой кнопкой мыши команду Перейти к объекту и выберите необходимое действие в пункте меню 'Противоречие решено?'.";

            SendHTMLMessage(subject, bodyText, mailRecipients, mailCopyRecipients, new List<Conflict> { conflict });
        }

        private void SendMessageAboutPerformerAnswer(Conflict conflict)
        {
            string otvet = conflict.Answer;

            //Добавляем адресатов сообщения
            List<UserReferenceObject> mailRecipients = new List<UserReferenceObject> { conflict.Declarant };
            string subject = "На созданное Вами противоречие дан ответ";
            string bodyText = string.Format("Ответ на <a href={0}>противоречие</a>: ", HTML_MailMessage.GetLinkFor(conflict.ReferenceObject)) + otvet;

            SendHTMLMessage(subject, bodyText, mailRecipients, new List<UserReferenceObject>(), new List<Conflict> { conflict });
        }

        private static void ChangeStage(Conflict conflict, string stageName)
        {
            Stage stage = conflict.ReferenceObject.Reference.ParameterGroup.Scheme.Stages.Where(s => s.Stage.Name == stageName).First().Stage;
            stage.Change(new List<ReferenceObject> { conflict.ReferenceObject });
        }

        private static string CreateHTMLBodyText(Conflict conflict)
        {
            string text = "В модуле \"Решение противоречий\", "
                                + conflict.Declarant.ToString()
                                + string.Format(", создал новое противоречие. Ознакомьтесь с противоречием, перейдя по <a href={0}>ссылке</a>, или перейдя на вкладку.", HTML_MailMessage.GetLinkFor(conflict.ReferenceObject));
            return text;
        }

        void SendHTMLMessage(string textSubj, string textMsg, List<UserReferenceObject> mailRecipients, List<UserReferenceObject> mailCopyRecipients, IEnumerable<Conflict> conflicts, List<ReferenceObject> attachments = null)
        {
            //исключаем из рассылки Острового
            UserReference userRef = new UserReference(ServerGateway.Connection);
            UserReferenceObject ostrovoy = userRef.Find(new Guid("e86c95d6-dadd-4839-b5ce-efea2e88b1a3")) as UserReferenceObject;

            string HTML_ErrandsTable = "";// CreateHTMLTableForErrands(errands);
            StringBuilder strBld = new StringBuilder();
            strBld.Append(String.Format(@"
    <p>Здравствуйте!</p>
    <p> {0} </p>
<p> {1} </p>", textMsg, HTML_ErrandsTable));

            //Добавление адресатов
            List<ReferenceObject> tempUsers = new List<ReferenceObject>();
            foreach (UserReferenceObject mailgroup in mailRecipients)
            {
                //if (mailgroup.ToString() == CurrentUserName) continue;
                if (mailgroup is User)
                {
                    tempUsers.Add(mailgroup);
                }
                List<User> internUsers = mailgroup.GetAllInternalUsers();
                if (internUsers != null) tempUsers.AddRange(mailgroup.GetAllInternalUsers());
            }

            List<ReferenceObject> users = new List<ReferenceObject>();
            foreach (var item in tempUsers)
            {
                users.Add(item);
                users.AddRange(item.Children);
            }
            if (users.Distinct() == null) return;
            //Исключаем из рассылки Острового
            if (ostrovoy != null && users.Remove(ostrovoy))
            {//проверяем чтобы остались адресаты в рассылке
                if (users.Count == 0) return;
            }
            List<string> externalEmails = new List<string>();

            string textMessage = strBld.ToString();
            List<ReferenceObject> attacments = new List<ReferenceObject>();
            foreach (var conflict in conflicts)
            {
                attacments.Add(conflict.ReferenceObject);
            }
            HTML_MailMessage message = new HTML_MailMessage(textSubj, strBld.ToString(), users, attacments);
            message.MailCopyRecipients = mailCopyRecipients.Select(u => u as ReferenceObject).ToList();
            message.Send();
        }


    }

    struct StageName
    {
        public const string NewConflict = "Новое противоречие";
        public const string CreationConflict = "Создание противоречие";
        public const string AnswerReceived = "Ответ на противоречие получен";
        public const string Resolved = "Вопрос закрыт";
        public const string NotResolved = "Вопрос не закрыт";
        public const string SentForResolving = "Направлено на решение";
        public const string HaveNoAnswer = "Ответ на противоречие не получен";
    }

    /// <summary>
    /// Противоречие
    /// </summary>
    class Conflict
    {
        public Conflict(ReferenceObject referenceObject)
        {
            this.ReferenceObject = referenceObject;
        }

        public ReferenceObject ReferenceObject { get; private set; }


        User _Declarant;

        /// <summary>
        /// Заявитель
        /// </summary>
        public User Declarant
        {
            get
            {
                if (_Declarant == null)
                {
                    _Declarant = this.ReferenceObject.SystemFields.Author;
                }
                return _Declarant;
            }
        }

        List<UserReferenceObject> _DeclarantManagers;

        /// <summary>
        /// Список руководителей заявителя
        /// </summary>
        public IEnumerable<UserReferenceObject> DeclarantManagers
        {
            get
            {
                if (_DeclarantManagers == null)
                {
                    _DeclarantManagers = this.ReferenceObject.GetObjects(declarantManagers_link_Guid).Select(r => r as UserReferenceObject).ToList();
                }
                return _DeclarantManagers;
            }
        }

        List<UserReferenceObject> _ExecutorManagers;

        /// <summary>
        /// Список руководителей исполнителя
        /// </summary>
        public IEnumerable<UserReferenceObject> ExecutorManagers
        {
            get
            {
                if (_ExecutorManagers == null)
                {
                    _ExecutorManagers = this.ReferenceObject.GetObjects(executorManagers_link_Guid).Select(r => r as UserReferenceObject).ToList();
                }
                return _ExecutorManagers;
            }
        }


        User _Performer;
        /// <summary>
        /// Исполнитель
        /// </summary>
        public User Performer
        {
            get
            {
                if (_Performer == null)
                {
                    ReferenceObject performer = this.ReferenceObject.GetObject(link_Performer_Guid);
                    if (performer != null) _Performer = performer as User;
                }
                return _Performer;
            }
        }

        public string Answer
        {
            get { return this.ReferenceObject[param_Answer_Guid].GetString(); }
        }

        public string Category
        {
            get { return this.ReferenceObject[category_param_Guid].GetString(); }
        }

        private readonly Guid param_Answer_Guid = new Guid("d99f6a48-7895-4dc9-aa5a-47d60158586c"); // Противоречие - поле Ответ
        private readonly Guid declarantManagers_link_Guid = new Guid("a8ced5e5-c8d4-4c70-a54c-fecbf4a959e1"); //связь N:N Руководитель заявителя (группы и пользователи)
        private readonly Guid executorManagers_link_Guid = new Guid("64ff13ea-2dac-4cc4-8def-e79340be0642"); //связь N:N Начальник отдела (с ИО) исполнителя
        private readonly Guid executorBoss_link_Guid = new Guid("ecb3497e-d186-4c7d-8ba1-a74f64734b2d"); //связь N:N Руководитель подразделения исполнителя
        private readonly Guid executorViceBoss_link_Guid = new Guid("0df30206-f00a-43fc-bd2d-5d21ce834dc2"); //связь N:N Заместитель Руководителя подразделения исполнителя
        private readonly Guid link_Performer_Guid = new Guid("755741dc-89ed-4e95-bc06-7dbe55032814"); //связь N:1 Противоречие - Исполнитель

        private readonly Guid responseText_param_Guid = new Guid("21ee9f6a-3b1c-4986-8475-0bc93b699c4e"); //параметр Отдел исполнителя
        private readonly Guid category_param_Guid = new Guid("2026f449-a670-4f60-9696-7dc0dbbe3a09"); //параметр Категория вопроса
    }

    /// <summary>
    /// HTML письмо
    /// ver 1.8
    /// </summary>
    class HTML_MailMessage
    {
        public HTML_MailMessage()
        {
            Subject = "";
            BodyText = "";
            MailAttachments = new List<ReferenceObject>();
            MailRecipients = new List<ReferenceObject>();
            MailCopyRecipients = new List<ReferenceObject>();
        }

        public HTML_MailMessage(string messageSubject, string messageText, List<ReferenceObject> mailRecipients, List<ReferenceObject> works)
        {
            this.Subject = messageSubject;
            this.BodyText = messageText;
            this.MailRecipients = mailRecipients;
            this.MailAttachments = works;
            this.MailCopyRecipients = new List<ReferenceObject>();
        }
        /*
        public static string GetHTMLTableSigns(ReferenceObject referenceObject)
        {
            StringBuilder table = new StringBuilder();
            table.Append(@"<table  cellpadding=""0"" cellspacing=""0"" border=""1"">

                 <tr>
                    <th nowrap bgcolor =#e1e1e1> № </th>
                    <th nowrap bgcolor =#e1e1e1> Тип подписи </th>
                    <th nowrap bgcolor =#e1e1e1> ФИО </th>
                    <th nowrap bgcolor =#e1e1e1> Дата подписи </th>
                    <th nowrap bgcolor =#e1e1e1> Резолюция </th>
                 </tr> ");
            int index = 1;
            var списокВсехПодписейОбъекта = referenceObject.Signatures;
            списокВсехПодписейОбъекта.GroupBy(s => s.SignatureObjectType);
            foreach (var sign in списокВсехПодписейОбъекта)
            {
                if (sign.Actual)
                {
                    string signDate = "";
                    if (sign.SignatureDate != null)
                        signDate = ((DateTime)sign.SignatureDate).ToShortDateString();
                    table.Append(String.Format(@"<tr>
                    <td> {0} </td>
                    <td> {1} </td>
                    <td> {2} </td>
                    <td> {3}&nbsp</td>
                    <td> {4}&nbsp</td>
                 </tr>", index++, sign.SignatureObjectType.Name, sign.UserName, signDate, sign.Resolution));
                }
                else continue;
            }
            table.Append("</table>");
            return table.ToString();
        }
        */
        public static string GetLinkFor(ReferenceObject object_ro)
        {
            if (object_ro == null) return null;
            return String.Format(@"docs://{0}/OpenReferenceWindow/?refId={1}&objId={2}",
                                 Connection.ConnectionParameters.GetServerAddress(), object_ro.Reference.Id, object_ro.SystemFields.Id);
        }
        public string Subject { get; set; }
        public string BodyText { get; set; }
        List<ReferenceObject> MailAttachments { get; set; }
        public bool AddAttachment(ReferenceObject attachment)
        {
            if (attachment == null) return false;
            if (MailAttachments == null) MailAttachments = new List<ReferenceObject>();
            MailAttachments.Add(attachment);
            return true;
        }

        List<string> _pathesToLocalAttachments;
        public void AddLocalFileByPath(string pathToFile)
        {
            if (_pathesToLocalAttachments == null) _pathesToLocalAttachments = new List<string>();
            _pathesToLocalAttachments.Add(pathToFile);
        }
        public List<ReferenceObject> MailRecipients { get; set; }
        List<ReferenceObject> _MailCopyRecipients;
        public List<ReferenceObject> MailCopyRecipients { get; set; }
        public void Send()
        {
            MailMessage = new MailMessage(Connection.Mail.DOCsAccount);
            MailMessage.Subject = CleanString(Subject);
            string messageBody = CreateMessageBody();
            MailMessage.SetBody(messageBody, MailBodyType.Html);
            FillAttachments();
            FillLocalFiles();
            FillMailRecipients();
            FillMailCopyRecipients();
            MailMessage.Send();
        }

#if client
        public void Show()
        {
            MailMessage = new MailMessage(Connection.Mail.DOCsAccount);
            MailMessage.Subject = CleanString(Subject);
            string messageBody = CreateMessageBody();
            MailMessage.SetBody(messageBody, MailBodyType.Html);
            FillAttachments();
            FillLocalFiles();
            FillMailRecipients();
            FillMailCopyRecipients();
            TFlex.DOCs.UI.DialogManager.Instance.ShowPropertyDialog(MailMessage);
        }
#endif
        private void FillLocalFiles()
        {
            if (_pathesToLocalAttachments == null) return;
            foreach (var localFilePath in _pathesToLocalAttachments)
            {
                MailMessage.Attachments.Add(new FileAttachment(localFilePath));
            }
        }

        private void FillMailRecipients()
        {
            //Добавление адресатов
            List<ReferenceObject> tempUsers = new List<ReferenceObject>();
            foreach (ReferenceObject recipient_ro in MailRecipients)
            {
                UserReferenceObject recipient_uro = recipient_ro as UserReferenceObject;
                if (recipient_uro is User)
                {
                    tempUsers.Add(recipient_ro);
                }
                else if (recipient_uro is UsersGroup)
                {
                    List<User> internUsers = recipient_uro.GetAllInternalUsers();
                    if (internUsers != null) tempUsers.AddRange(recipient_uro.GetAllInternalUsers());
                }
                else continue;
            }
            List<ReferenceObject> users = new List<ReferenceObject>();
            foreach (var item in tempUsers)
            {
                users.Add(item);
                users.AddRange(item.Children);
            }
            if (users.Distinct() == null) return;

            foreach (UserReferenceObject user in users.Distinct())
            {
                User mailuser = user as User;
                MailMessage.To.Add(new MailUser(mailuser));

                if (mailuser.Email != null && !mailuser.Email.IsEmpty)
                {
                    MailMessage.To.Add(new EMailAddress(mailuser.Email.ToString()));
                }
            }
        }
        private void FillMailCopyRecipients()
        {
            //Добавление адресатов
            List<ReferenceObject> tempUsers = new List<ReferenceObject>();
            foreach (ReferenceObject recipient_ro in MailCopyRecipients)
            {
                UserReferenceObject recipient_uro = recipient_ro as UserReferenceObject;
                if (recipient_uro is User)
                {
                    tempUsers.Add(recipient_ro);
                }
                else if (recipient_uro is UsersGroup)
                {
                    List<User> internUsers = recipient_uro.GetAllInternalUsers();
                    if (internUsers != null) tempUsers.AddRange(recipient_uro.GetAllInternalUsers());
                }
                else continue;
            }
            List<ReferenceObject> users = new List<ReferenceObject>();
            foreach (var item in tempUsers)
            {
                users.Add(item);
                users.AddRange(item.Children);
            }
            if (users.Distinct() == null) return;

            foreach (UserReferenceObject user in users.Distinct())
            {
                User mailuser = user as User;
                MailMessage.Copy.Add(new MailUser(mailuser));

                if (mailuser.Email != null && !mailuser.Email.IsEmpty)
                {
                    MailMessage.Copy.Add(new EMailAddress(mailuser.Email.ToString()));
                }
            }
        }


        private void FillAttachments()
        {
            foreach (var attachment in MailAttachments)
            {
                Guid fileClassGuid = new Guid("4731e1b6-b27e-4895-be2f-b8140316bfc0"); //гуид абстрактного класса Файл справочника Файлы
                if (attachment.Class.IsInherit(fileClassGuid))
                {
                    FileObject file = (FileObject)attachment;
                    file.GetHeadRevision();
                    FileAttachment fileAttachment = new FileAttachment(file.LocalPath);
                    fileAttachment.DownloadFile();
                    MailMessage.Attachments.Add(fileAttachment);
                }
                else MailMessage.Attachments.Add(new ObjectAttachment(attachment));
            }
        }

        MailMessage MailMessage;
        string BodyHeaderForSystem =
            @"<p>Здравствуйте!</p>
    <p>Это автоматически созданное письмо, не отвечайте на него.</p>";

        string BodyFooterForSystem { get { return String.Format(@"<p class=MsoNormal>
<span style='font-size:12.0pt;font-family:""Arial"",""sans-serif"";mso-fareast-font-family:
""Times New Roman"";mso-fareast-theme-font:minor-fareast;mso-fareast-language:
RU;mso-no-proof:yes'>Если вы получили это письмо ошибочно или у вас возникли вопросы обращайтесь в
<a href=""mailto:Отдел%20автоматизации"">Отдел автоматизации</a>
 по тел. 1180, 1200, 1204<o:p></o:p></p>
<p>или по E-mail: <a href=""mailto:Отдел%20автоматизации"">Отдел автоматизации</a></span>
<br>
<img src=""data:image/png;base64,{0}  alt=""АСУ ""Динамика"""">
</p>", Logo_Base64); } }

        private string CreateMessageBody()
        {
            string bodyHeader = "";
            string bodyFooter = "";
            string messageBody = "";
            if (FromSystem)
            {
                bodyHeader = BodyHeaderForSystem;
                bodyFooter = BodyFooterForSystem;
            }
            else
            {
                bodyFooter = String.Format(@"
<p> C уважением,
<br>{0}.
<br>
<img src=""data:image/png;base64,{1}  alt=""АСУ ""Динамика"""">
</p>", Connection.ClientView.UserName, Logo_Base64);
                //bodyFooter = "";

            }

            messageBody = String.Format(
                @"
<html>
    <head>
		<meta charset=""utf-8"">
	</head>
    <body>
        <span style='font-size:12.0pt; font-family:""Arial"",""sans-serif"";mso-fareast-font-family:""Times New Roman"";
            mso-fareast-theme-font:minor-fareast;mso-fareast-language:RU;mso-no-proof:yes'>
{0}
          <p> 
{1} 
          </p>
{2}
        </span>
    </body>
</html>", bodyHeader, BodyText, bodyFooter);

            return messageBody;
        }

        public bool FromSystem { get { return Connection.ClientView.GetUser().IsSystem; } }
        static public string CleanString(string s)
        {
            if (s != null && s.Length > 0)
            {
                StringBuilder sb = new StringBuilder(s.Length);
                foreach (char c in s)
                {
                    if (Char.IsControl(c) == true)
                        continue;
                    sb.Append(c);
                }
                s = sb.ToString();
            }
            return s;
        }

        const int LogoFileID = 68387;

        string _Logo_Base64;

        private string Logo_Base64
        {
            get
            {
                if (_Logo_Base64 == null)
                    _Logo_Base64 = GetDecodedImage();
                return _Logo_Base64;
            }
        }

        private string GetDecodedImage()
        {
            FileReference fileReference = new FileReference(Connection);
            ReferenceObject logoFile_ro = fileReference.Find(LogoFileID);
            if (logoFile_ro == null) return "";
            FileObject logoFile = logoFile_ro as FileObject;
            logoFile.GetHeadRevision();
            FileStream FStream = new FileStream(logoFile.LocalPath, FileMode.Open, FileAccess.Read);
            // Создаем BinaryReader
            BinaryReader sr = new BinaryReader(FStream);
            byte[] byteArray;
            // Пока не достигнут конец файла считываем его побайтно
            using (BinaryReader br = new BinaryReader(FStream))
            {
                byteArray = br.ReadBytes((int)FStream.Length);
            }
            sr.Close();

            String DecodedImage = System.Convert.ToBase64String(byteArray);
            return DecodedImage;
        }

        static ServerConnection Connection { get { return ServerGateway.Connection; } }
    }
}