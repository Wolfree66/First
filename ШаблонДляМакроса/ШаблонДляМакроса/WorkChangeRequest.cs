//#define server
//#define test
//#define serverTest
//Исходными даными являются записи из журнала редактирования, при создании записи автоматически генерируется запрос об изменении работы
//т.к. записи создаются в подразделениях, то в запросе имеются данные о работе из плана подразделения и типе необходимой коррекции
namespace WorkChangeRequest
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Windows.Forms;
    using TFlex.DOCs.Model;
    using TFlex.DOCs.Model.Classes;
    using TFlex.DOCs.Model.Macros;

    using TFlex.DOCs.Model.References;

    using TFlex.DOCs.Model.References.Files;
    using System.IO;
    using TFlex.DOCs.Model.References.Users;
    using TFlex.DOCs.Model.Mail;
    using TFlex.DOCs.Model.Search;
    using TFlex.DOCs.Model.Structure;
    using TFlex.DOCs.Model.Stages;
    using TFlex.DOCs.Model.Parameters;

#if serverTest
    using Excel = NPOI.SS.UserModel;
#endif


#if !server
    using TFlex.DOCs.Model.Macros.ObjectModel;
    using TFlex.DOCs.UI.Objects.Managers;
    using TFlex.DOCs.UI.Common;
    //using TFlex.DOCs.Model.Types;
#endif

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
#if test
        public void Test()
        {
            //проверка работы в архиве
            //ReferenceObject work = References.ProjectManagementReference.Find(363684);
            //ReferenceObject work = References.ProjectManagementReference.Find(363658);
            //ProjectManagementWork pmWork = Factory.Create_ProjectManagementWork(work);
            //bool isArchieved = Is_Work_Archieved(pmWork);
            //return;

            //ЗакрытьВсеНезакрытыеДублёры();
            ReferenceObject first = References.WorkChangeRequestReference.Find(2964);
            WorkChangeRequest firstWorkChangeRequest = Factory.Create_WorkChangeRequest(first); //создаём новый запрос
            ReferenceObject second = References.WorkChangeRequestReference.Find(7989);
            WorkChangeRequest secondWorkChangeRequest = Factory.Create_WorkChangeRequest(second); //создаём новый запрос
            ReferenceObject third = References.WorkChangeRequestReference.Find(7990);
            WorkChangeRequest thirdWorkChangeRequest = Factory.Create_WorkChangeRequest(second); //создаём новый запрос

            CheckRequestStatus(firstWorkChangeRequest);
            return;
            Dictionary<WorkChangeRequest, string> errors = new Dictionary<WorkChangeRequest, string>();
            errors.Add(firstWorkChangeRequest, "Error 1");
            errors.Add(secondWorkChangeRequest, "Error 2");
            errors.Add(thirdWorkChangeRequest, "Error 3");
            SendMessageToPlanningDepartmentsAboutProblem(errors);
            //            string excelFilePath = CreateExcelTable(errors);
            //CheckRequestStatus(firstWorkChangeRequest);
            //SendMessageNeedManualClose(firstWorkChangeRequest);
            //firstWorkChangeRequest.Close(WorkChangeRequest.CompleteNote.NewRequestCreation);
        
        }
#endif
        List<string> Columns = new List<string> {
            "№",
            "Дата создания",
            "Тип запроса",
            "Содержание запроса",
            "Название работы",
            "Исполнитель",
            "Номер проекта",
            "ФИО руководителя проекта",
            "План подразделения",
            "Решение РП",
            "Причина отклонения",
            "Комментарий"};


        #region Вычисляемые колонки

        public string ДатаНачалаСвязаннойРаботы()
        {
            ReferenceObject current = Context.ReferenceObject;
            WorkChangeRequest workChangeRequest = Factory.Create_WorkChangeRequest(current);
            //если запрос не на изменение даты то возвращаем пустую строку
            if (workChangeRequest as WorkChangeRequest_DatesChange == null) return "";
            ProjectManagementWork departmentWork = workChangeRequest.DepartmentWork;
            string startDate = "Не доступно";
            if (!departmentWork.IsEmpty) startDate = departmentWork.StartDate.ToShortDateString();
            return startDate;
        }
        public string ДатаОкончанияСвязаннойРаботы()
        {
            ReferenceObject current = Context.ReferenceObject;
            WorkChangeRequest workChangeRequest = Factory.Create_WorkChangeRequest(current);
            //если запрос не на изменение даты то возвращаем пустую строку
            if (workChangeRequest as WorkChangeRequest_DatesChange == null) return "";
            ProjectManagementWork departmentWork = workChangeRequest.DepartmentWork;
            string endDate = "Не доступно";
            if (!departmentWork.IsEmpty) endDate = departmentWork.EndDate.ToShortDateString();
            return endDate;
        }
        public string ДатаНачалаСинхронизированнойРаботы()
        {
            ReferenceObject current = Context.ReferenceObject;
            WorkChangeRequest workChangeRequest = Factory.Create_WorkChangeRequest(current);
            //если запрос не на изменение даты то возвращаем пустую строку
            if (workChangeRequest as WorkChangeRequest_DatesChange == null) return "";
            ProjectManagementWork work = workChangeRequest.GetManagerWork();
            string result = "Не доступно";
            if (work != null && !work.IsEmpty) result = work.StartDate.ToShortDateString();
            return result;
        }
        public string ДатаОкончанияСинхронизированнойРаботы()
        {
            ReferenceObject current = Context.ReferenceObject;
            WorkChangeRequest workChangeRequest = Factory.Create_WorkChangeRequest(current);
            //если запрос не на изменение даты то возвращаем пустую строку
            if (workChangeRequest as WorkChangeRequest_DatesChange == null) return "";
            ProjectManagementWork work = workChangeRequest.GetManagerWork();
            string result = "Не доступно";
            if (work != null && !work.IsEmpty) result = work.EndDate.ToShortDateString();
            return result;
        }
        #endregion

        #region Административные действия

        public void ВернутьНаОбработанРП()
        {
            ReferenceObject current = Context.ReferenceObject;
            WorkChangeRequest workChangeRequest = Factory.Create_WorkChangeRequest(current); //создаём новый запрос
            workChangeRequest.CloseDate = null;
            workChangeRequest.CommentToComplete = "";
            workChangeRequest.SaveToDB();
            workChangeRequest.ChangeStageTo(WorkChangeRequest.Stages.DoneByManager);

        }

        public void ВернутьНаНовый()
        {
            ReferenceObject current = Context.ReferenceObject;
            WorkChangeRequest workChangeRequest = Factory.Create_WorkChangeRequest(current); //создаём новый запрос
            workChangeRequest.CloseDate = null;
            workChangeRequest.CommentToComplete = "";
            workChangeRequest.CompletedByManagerDate = null;
            workChangeRequest.ManagerDecision = "";
            workChangeRequest.SaveToDB();
            workChangeRequest.ChangeStageTo(WorkChangeRequest.Stages.New);

        }
        #endregion


        #region Входные События

        public void ЗакрытьВсеНезакрытыеДублёры()
        {
            var requests = GetNotCompletedRequests().OrderByDescending(r => r.CreationDate);
            List<Guid> workGuids = new List<Guid>();
            List<WorkChangeRequest> toClose = new List<WorkChangeRequest>();
            foreach (var item in requests)
            {
                if (workGuids.Contains(item.DepartmentWorkGUID))
                {
                    int id = item.ReferenceObject.SystemFields.Id;
                    toClose.Add(item);
                }
                else workGuids.Add(item.DepartmentWorkGUID);
            }

            foreach (var item in toClose)
            {
                item.Close(WorkChangeRequest.CompleteNote.ManualClose);
            }

        }

        /// <summary>
        /// Закрывает текущий запрос
        /// </summary>
        public void ЗакрытьТекущийЗапрос()
        {
            WorkChangeRequest newWorkChangeRequest = Factory.Create_WorkChangeRequest(Context.ReferenceObject); //создаём новый запрос

            newWorkChangeRequest.Close(WorkChangeRequest.CompleteNote.ManualClose);
        }

        /// <summary>
        /// Метод создающий запрос при завершении сохранения записи из журнала редактирования
        /// Обрабатывается на стороне сервера
        /// </summary>
        public void Событие_ЗавершениеСохраненияЗаписиВЖурналеРедактирования()
        {
            ReferenceObject editReasonNote_ro = Context.ReferenceObject;

            EditReasonNote editReasonNote = new EditReasonNote(editReasonNote_ro); //получаем запись журнала редактирования
            //пишется запись в журнал редактирования и удаляется работа поэтому, если работы нет, то и запись не обрабатывается
            if (editReasonNote.Work.IsEmpty) return;// throw new Exception(string.Format("Не найдена работа для записи Журнала Редактирования, Reference - {0}, ObjectGuid - {1}", editReasonNote.Reference.Name, editReasonNote.Guid.ToString()));

            if (WorkNotFromDepartmentPlan(editReasonNote.Work)) //если связанная работа не из плана подразделения, то завершаем метод


            { return; }

            if (editReasonNote.TypeCode == (int)EditReasonNote.Type.NotDefined) //если она типа "Не определено" то закрываем все незакрытые запросы по связанной работе
            {
                CloseAllUncompletedChangeRequestsForWork(editReasonNote.Work);
            }
            else if (IsNeedToCreateChangeRequest(editReasonNote))//если удовлетворяются требования для создания запроса, создаём новый запрос
            {
                WorkChangeRequest newWorkChangeRequest = Factory.Create_WorkChangeRequest(editReasonNote); //создаём новый запрос
                //закрываем все ранее созданные запросы
                var requests = GetUncompletedChangeRequestsForWork(editReasonNote.Work);
                foreach (var item in requests)
                {
                    if (item.ReferenceObject.SystemFields.Id == newWorkChangeRequest.ReferenceObject.SystemFields.Id) continue;
                    item.Close(WorkChangeRequest.CompleteNote.NewRequestCreation);
                }

                //при создании нового запроса сообщение не возникает
                //SendMessageAboutNewRequest(newWorkChangeRequest);
            }
        }

        /// <summary>
        /// Обрабатывается на стороне сервера
        /// </summary>
        public void Событие_ИзменениеДатыОбработкиРП()
        {

            ReferenceObject current = Context.ReferenceObject;
            WorkChangeRequest workChangeRequest = Factory.Create_WorkChangeRequest(current);
            try
            {
                workChangeRequest.ChangeStageTo(WorkChangeRequest.Stages.DoneByManager);//меняем стадию на "Обработано РП"
                if (!workChangeRequest.TryAutomaticallyDoCloseActions())//пытаемся выполнить действия необходимые для закрытия запроса
                {
                    //если не удалось выполнить
                    SendMessageNeedManualClose(workChangeRequest);//посылаем письмо о необходимости ручного закрытия запроса
                }
                else //если все действия выполнены
                {
                    if (workChangeRequest.CanChangeStageForAcceptedRequestToClose) //пытаемся изменить стадию на "Закрыт"
                    {
                        workChangeRequest.Close(WorkChangeRequest.CompleteNote.AutoClose);

                    }
                    else SendMessageNeedManualClose(workChangeRequest);//посылаем письмо о необходимости ручного закрытия запроса
                }
            }
            catch (Exception e)
            {
                throw new Exception(string.Format("Событие_ИзменениеДатыОбработкиРП, Reference - {0}, ObjectGuid - {1}", workChangeRequest.ReferenceObject.Reference.Name, workChangeRequest.ReferenceObject.SystemFields.Guid.ToString()) + e.ToString());
            }
        }

        /// <summary>
        /// Обрабатывается на стороне сервера
        /// </summary>
        public void Событие_ИзменениеДатыЗакрытияЗапроса()
        {
            ReferenceObject current = Context.ReferenceObject;
            WorkChangeRequest workChangeRequest = Factory.Create_WorkChangeRequest(current);

            workChangeRequest.ChangeStageTo(WorkChangeRequest.Stages.Completed);
            //информируются заинтересованные лица
            //при Закрытии сообщение не возникает
            //SendMessageAboutClose(workChangeRequest);
        }

#if !server
        /// <summary>
        /// Вызывается РП
        /// </summary>
        public void Событие_ПрименитьРешение()
        {
            ReferenceObject current = Context.ReferenceObject;
            WorkChangeRequest workChangeRequest = Factory.Create_WorkChangeRequest(current);
            if (IsCurrentUserHaveRole(workChangeRequest, WorkChangeRequest.Role.Manager) || ((User)CurrentUser).Class.IsAdministrator)
            {
                if (workChangeRequest.CompletedByManagerDate != null) //проверяем было ли уже применено решение РП
                {
                    MessageBox.Show("Запрос уже обработан руководителем проекта!");
                    return;
                }

                //для изменения исполнителя необходимо задать нового исполнителя
                WorkChangeRequest_ChangeResource changeResourceRequest = workChangeRequest as WorkChangeRequest_ChangeResource;
                if (changeResourceRequest != null && !ChooseNewResource(changeResourceRequest))
                { return; }

                if (workChangeRequest.DoAcceptActions())//выполняем необходимые действия по обработке запроса
                {
                    if (workChangeRequest.IsChangeStageCanBeComplete)
                        workChangeRequest.CompleteByManager("");// переводим в обработано РП
                }
            }
            else { MessageBox.Show("У вас недостаточно прав для выполнения операции!", "Предупреждение!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }

        }



        /// <summary>
        /// Вызывается РП
        /// </summary>
        public void Событие_ОтклонитьЗапрос()
        {
            ReferenceObject current = Context.ReferenceObject;
            WorkChangeRequest workChangeRequest = Factory.Create_WorkChangeRequest(current);
            if (IsCurrentUserHaveRole(workChangeRequest, WorkChangeRequest.Role.Manager) || ((User)CurrentUser).Class.IsAdministrator)
            {
                if (workChangeRequest.CompletedByManagerDate != null) //проверяем было ли уже применено решение РП
                {
                    MessageBox.Show("Запрос уже обработан руководителем проекта!");
                    return;
                }
                //запрашиваем причину отклонения
                ДиалогВвода диалог = СоздатьДиалогВвода("Укажите причину отклонения запроса.");
                диалог.ДобавитьСтроковое("Причина", "", true, true);
                string reason = "";
                if (диалог.Показать())
                {
                    reason = диалог["Причина"];

                    if (!string.IsNullOrWhiteSpace(reason))
                    {
                        string changeTypeName = workChangeRequest.Type;

                        workChangeRequest.RejectByManager(reason, "");
                        //SendMessageAboutReject(workChangeRequest);
                    }
                    else
                    {
                        MessageBox.Show("Необходимо указать причину отклонения запроса!");
                    }
                }

            }
            else { MessageBox.Show("У вас недостаточно прав для выполнения операции!", "Предупреждение!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
        }

        /// <summary>
        /// Вызывается работником ПлОП
        /// </summary>
        public void Событие_ЗакрытьЗапрос()
        {
            ReferenceObject current = Context.ReferenceObject;
            WorkChangeRequest workChangeRequest = Factory.Create_WorkChangeRequest(current);

            if (IsCurrentUserHaveRole(workChangeRequest, WorkChangeRequest.Role.DepartmentScheduler) || ((User)CurrentUser).Class.IsAdministrator)//текущий пользователь имеет доступ к плану подразделения
            {
                if (workChangeRequest.CloseDate != null) //проверяем было ли уже закрыто ПлОП
                {
                    MessageBox.Show("Запрос уже закрыт!");
                    return;
                }
                if (workChangeRequest.ManagerDecision == WorkChangeRequest.ManagerDecisionValue.Reject)
                {
                    //запрашиваем нужно ли создавать новый запрос на основе старого
                    DialogResult dialogResult = MessageBox.Show("Создать новый запрос на основе закрываемого?", "Создать новый запрос?", MessageBoxButtons.YesNoCancel);
                    if (dialogResult == DialogResult.No)
                    {
                        workChangeRequest.Close("Закрыто в подразделении.");
                    }
                    else if (dialogResult == DialogResult.Yes)
                    {
                        EditReasonNote editReasonNote = new EditReasonNote(workChangeRequest.EditReasonUID);
                        WorkChangeRequest newWorkChangeRequest = Factory.Create_WorkChangeRequest(editReasonNote); //создаём новый запрос
                        workChangeRequest.Close("Закрыто в подразделении.");

                        //при создании нового запроса сообщение не возникает
                        //SendMessageAboutNewRequest(newWorkChangeRequest);
                    }
                }
                else if (workChangeRequest.ManagerDecision == WorkChangeRequest.ManagerDecisionValue.Accept)
                {
                    workChangeRequest.TryAutomaticallyDoCloseActions();
                    workChangeRequest.Close("Закрыто в подразделении.");
                }
            }
            else { MessageBox.Show("У вас недостаточно прав для выполнения операции!", "Предупреждение!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
        }

        public void Событие_НажатиеНаКнопкуПерейтиКРаботеВПланРП()
        {
            ReferenceObject current = Context.ReferenceObject;
            WorkChangeRequest workChangeRequest = Factory.Create_WorkChangeRequest(current);
            ProjectManagementWork managerWork = workChangeRequest.GetManagerWork();

            if (managerWork != null && !managerWork.IsEmpty)
            {
                var ro = managerWork.ReferenceObject;

                string path = System.IO.Path.Combine(Application.StartupPath, "TFlex.DOCs.UI.ProjectManagement.dll"); // Здесь задается путь до библиотеки.
                System.Reflection.Assembly assembly = System.Reflection.Assembly.LoadFile(path);
                var type = assembly.GetType("TFlex.DOCs.UI.Controls.ProjectEnvironmentManager");
                var method = type.GetMethods().Where(m => m.Name == "OpenProject" && m.GetParameters().Length == 4).First();

                //ReferenceObject work = null; // Работа
                string pathToFocusedObject = GetPathToRootObject(managerWork.ReferenceObject, managerWork.Project.ReferenceObject);

                method.Invoke(null, new object[] { managerWork.Project.ReferenceObject.SystemFields.Id, pathToFocusedObject, null, false });


            }
            else
            {
                MessageBox.Show("Не найдена синхронизированная работа в пространстве Рабочие проекты!", "Работа не найдена", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetPathToRootObject(ReferenceObject source, ReferenceObject root)
        {
            List<int> ids = new List<int>();
            var parentObject = source;
            while (parentObject != null)
            {
                ids.Add(parentObject.SystemFields.Id);
                if (root.SystemFields.Id == parentObject.SystemFields.Id)
                    break;
                parentObject = parentObject.Parent;
            }
            if (ids.Count == 0)
                return string.Empty;

            return string.Join("&", ids);
        }


        public void Событие_НажатиеНаКнопкуПерейтиКРаботеВСправочник()
        {
            ReferenceObject current = Context.ReferenceObject;
            WorkChangeRequest workChangeRequest = Factory.Create_WorkChangeRequest(current);
            ProjectManagementWork managerWork = workChangeRequest.GetManagerWork(true);
            //MessageBox.Show("1");

            if (managerWork != null && !managerWork.IsEmpty)
            {
                var ro = managerWork.ReferenceObject;

                UIMacroContext uiContext = Context as UIMacroContext;
                IApplicationFormView application = uiContext.FindCurrentApplicationFormView();
                application.OpenReferenceWindow(
                    ro.Reference.ParameterGroup.ReferenceInfo,
                    0,
                    ro.SystemFields.Id,
                    ro.Reference.PrototypeMode);
            }
            else
            {
                MessageBox.Show("Не найдена синхронизированная работа в пространстве Рабочие проекты!", "Работа не найдена", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
#endif

        #endregion

        #region Команды

        public void ПроверитьНаличиеЗапросаУРаботы()//неправильно работает
        {
            //нужно проверить возможность существования записи в ЖР и нескольких запросов коррекции 
            ReferenceObject work_ro = Context.ReferenceObject;
            ProjectManagementWork work = Factory.Create_ProjectManagementWork(work_ro);
            if (WorkNotFromDepartmentPlan(work)) //если связанная работа не из плана подразделения, то завершаем метод
            { return; }
            EditReasonNote editReasonNote = GetLastEditReasonNote(work);
            if (editReasonNote.TypeCode == (int)EditReasonNote.Type.NotDefined) return;
            if (IsNeedToCreateChangeRequest(editReasonNote))
            {
                var workCR = GetChangeRequests(work, editReasonNote, null).FirstOrDefault();
                if (workCR == null) workCR = Factory.Create_WorkChangeRequest(editReasonNote);
            }
        }
        #endregion

#if !server
        private bool ChooseNewResource(WorkChangeRequest_ChangeResource request)
        {

            //MacroContext mc = new MacroContext(References.Connection);
            UIMacroContext uIMacroContext = Context as UIMacroContext;
            var dialog = uIMacroContext.CreateInputDialog();
            //var dialog = ObjectCreator.CreateObject<IInputDialog>();
            //InputDialog dialog = new InputDialog(uIMacroContext, "Выберите нового исполнителя");
            dialog.Caption = "Выберите нового исполнителя";
            //Filter filter = null;
            dialog.AddSelectFromReferenceField("Выберите исполнителя", "Ресурсы", "", request.NewResource, true, "");
            dialog.AddIntegerField("Трудозатраты", request.NewLabourHours, true);

            if (dialog.Show(uIMacroContext))
            {
                request.NewResource = (ReferenceObject)dialog.GetValue("Выберите исполнителя");
                request.NewLabourHours = dialog.GetValue("Трудозатраты");
                return true;
            }
            return false;

        }


#endif

        private void SendMessageToPlanningDepartmentsAboutProblem(Dictionary<WorkChangeRequest, string> dictWCRsWithMissedManagersWork)
        {
#if serverTest
            //создаём словарь где ключ - проект подразделения, значение - список с запросами по этому проекту
            Dictionary<string, List<WorkChangeRequest>> dictProjectCRs = new Dictionary<string, List<WorkChangeRequest>>();
            foreach (var kvp in dictWCRsWithMissedManagersWork)
            {
                string project = kvp.Key.DepartmentProjectName;
                List<WorkChangeRequest> listCR = null;
                if (dictProjectCRs.TryGetValue(project, out listCR))
                {
                    listCR.Add(kvp.Key);
                }
                else
                {
                    listCR = new List<WorkChangeRequest> { kvp.Key };
                    dictProjectCRs.Add(project, listCR);
                }
            }
            //для каждого проекта 
            foreach (var item in dictProjectCRs)
            {
                var mailRecipients = item.Value.First().DepartmentWork.Project.GetUsersHaveAccess(ProjectManagementWork.AccessCode.NoLimit);
                if (mailRecipients.Count == 0) continue;

                Dictionary<WorkChangeRequest, string> dictErrors = new Dictionary<WorkChangeRequest, string>();
                foreach (var request in item.Value)
                {
                    dictErrors.Add(request, dictWCRsWithMissedManagersWork[request]);
                }

                //создаём таблицу Excel
                string failExcelTablePath = CreateExcelTable(dictErrors);
                //отправляем письмо с таблицей
                HTML_MailMessage message = new HTML_MailMessage();
                message.Subject = "Запросы коррекции у которых отсутствует синхронизация.";
                message.BodyText = "Во вложении файл, с запросами коррекции у которых отсутствует синхронизация.";
                message.AddLocalFileByPath(failExcelTablePath);
#if test
                message.MailRecipients = new List<ReferenceObject> { References.Connection.ClientView.GetUser() };
#else
                message.MailRecipients = mailRecipients.Select(u => u as ReferenceObject).ToList();
#endif
                message.Send();
                //удаляем файл с таблицей

                File.Delete(failExcelTablePath);
            }
#endif
        }

        private void SendMessageToAdminAboutProblem(Dictionary<WorkChangeRequest, string> dictWorkChangeRequestsProblem)
        {
            List<ReferenceObject> mailRecipients = new List<ReferenceObject> { References.UsersReference.FindUser("Липаев Алексей Александрович") };
            if (mailRecipients != null && mailRecipients.Count > 0)
            {
                string text = "Проблемные запросы <br>" + CreateHTMLTableForWorkChangeRequest(dictWorkChangeRequestsProblem);
                string subject = "Проблемные запросы";
                HTML_MailMessage message = new HTML_MailMessage(subject, text, mailRecipients, dictWorkChangeRequestsProblem.Keys.Select(wrc => wrc.ReferenceObject).ToList());
                message.Send();
            }
            //throw new NotImplementedException();
        }

        string TempFolderPath
        {
            get
            {
                //string path = ApplicationFormManager.Instance.MainForm.ReportFolderPath; // если задана папка для отчётов
                //if (String.IsNullOrEmpty(path)) path = Environment.GetFolderPath(Environment.SpecialFolder.Personal); //папка - Мои документы
                //string path = Environment.GetFolderPath(Environment.SpecialFolder.Personal); //папка - Мои документы
                return System.IO.Path.GetTempPath();//
            }
        }

#if serverTest
        #region NPOI library
        string CreateExcelTable(Dictionary<WorkChangeRequest, string> dictWorkChangeRequestsProblem)
        {
            string filePath = TempFolderPath + "\\ActionNeeds.xls";

            //NPOI.SS.UserModel.IWorkbook iWB = new NPOI.SS.UserModel.IWorkbook();
            NPOI.HSSF.UserModel.HSSFWorkbook wb = new NPOI.HSSF.UserModel.HSSFWorkbook();
            Excel.ISheet sheet = wb.CreateSheet();
            int numRow = 0;
            Excel.IRow rowHeader = sheet.CreateRow(numRow);
            numRow++;
            int numCell = 0;

            Excel.ICellStyle commonCellStyle = wb.CreateCellStyle();
            commonCellStyle.BorderBottom = Excel.BorderStyle.Thin;
            commonCellStyle.BorderLeft = Excel.BorderStyle.Thin;
            commonCellStyle.BorderRight = Excel.BorderStyle.Thin;
            commonCellStyle.BorderTop = Excel.BorderStyle.Thin;
            commonCellStyle.WrapText = true;
            commonCellStyle.VerticalAlignment = Excel.VerticalAlignment.Center;

            NPOI.SS.UserModel.ICellStyle headerCellStyle = wb.CreateCellStyle();
            headerCellStyle.CloneStyleFrom(commonCellStyle);
            headerCellStyle.Alignment = Excel.HorizontalAlignment.Center;
            headerCellStyle.VerticalAlignment = Excel.VerticalAlignment.Center;
            //headerCellStyle.ShrinkToFit = true;
            headerCellStyle.FillForegroundColor = 3;

            headerCellStyle.FillPattern = Excel.FillPattern.SolidForeground;
            var font = wb.CreateFont();
            font.Boldweight = 10;
            headerCellStyle.SetFont(font);
            //MessageBox.Show(cellStyle.FillBackgroundColorColor.Indexed.ToString());

            headerCellStyle.WrapText = true;

            //заполняем заголовок таблицы
            foreach (var columnName in Columns)
            {
                Excel.ICell headerCell = rowHeader.CreateCell(numCell, NPOI.SS.UserModel.CellType.String);
                headerCell.SetCellValue(columnName);
                headerCell.CellStyle = headerCellStyle;

                numCell++;
            }

            //создаём стиль для гиперссылок
            Excel.ICellStyle hyperLinkCellStyle = wb.CreateCellStyle();//new NPOI.HSSF.UserModel.HSSFCellStyle(2500, extendedFormatRecord, wb);
            hyperLinkCellStyle.CloneStyleFrom(commonCellStyle);
            Excel.IFont hyperLinkFont = wb.CreateFont();
            hyperLinkFont.Color = 4;
            hyperLinkFont.FontHeight = 240;
            hyperLinkFont.Underline = Excel.FontUnderlineType.Single;
            hyperLinkCellStyle.SetFont(hyperLinkFont);


            int counterCR = 1;
            foreach (var kvp in dictWorkChangeRequestsProblem)
            {
                WorkChangeRequest workChangeRequest = kvp.Key;
                //заполняем ячейки таблицы
                Excel.IRow row = sheet.CreateRow(numRow);
                //№
                AddCellToRow(row, counterCR, commonCellStyle);

                //Дата создания
                AddCellToRow(row, workChangeRequest.CreationDate.ToShortDateString(), commonCellStyle);

                //"Тип запроса",
                AddCellToRow(row, workChangeRequest.Type, commonCellStyle);

                //"Содержание запроса",
                Excel.ICell cellContent = AddCellToRow(row, workChangeRequest.Content, hyperLinkCellStyle);
                var creationHelper = new NPOI.HSSF.UserModel.HSSFCreationHelper(wb);
                var hyperLinkContent = creationHelper.CreateHyperlink(Excel.HyperlinkType.Url);
                hyperLinkContent.Address = HTML_MailMessage.GetLinkFor(workChangeRequest.ReferenceObject); ;
                cellContent.Hyperlink = hyperLinkContent;

                //"Название работы",
                string workName = "Не доступно";
                string workLink = "";
                ProjectManagementWork departmentWork = workChangeRequest.DepartmentWork;
                if (!departmentWork.IsEmpty)
                {
                    workName = departmentWork.Name;
                    workLink = HTML_MailMessage.GetLinkFor(departmentWork.ReferenceObject);




                }



                Excel.ICell cell = AddCellToRow(row, workName, hyperLinkCellStyle);

                var hyperLink = creationHelper.CreateHyperlink(Excel.HyperlinkType.Url);
                hyperLink.Address = workLink;
                cell.Hyperlink = hyperLink;
                //Исполнитель
                AddCellToRow(row, "В разработке", commonCellStyle);

                //"Номер проекта",
                AddCellToRow(row, workChangeRequest.ProjectNumberFrom1C, commonCellStyle);

                //"ФИО руководителя проекта",
                AddCellToRow(row, workChangeRequest.ManagerName, commonCellStyle);

                //"План подразделения",
                AddCellToRow(row, workChangeRequest.DepartmentProjectName, commonCellStyle);

                //"Решение РП",
                AddCellToRow(row, workChangeRequest.ManagerDecision, commonCellStyle);

                //"Причина отклонения",
                AddCellToRow(row, workChangeRequest.ReasonToReject, commonCellStyle);

                //"Комментарий"
                AddCellToRow(row, workChangeRequest.CommentToComplete, commonCellStyle);

                counterCR++;
                numRow++;
            }
            for (int i = 0; i < Columns.Count; i++)
            {
                sheet.AutoSizeColumn(i);
            }
            using (FileStream fileStream = new FileStream(filePath, FileMode.Create))
            {
                wb.Write(fileStream);



            }
            return filePath;
        }

        private static Excel.ICell AddCellToRow(Excel.IRow row, int cellValue, Excel.ICellStyle cellstyle)
        {
            Excel.ICell newCell = row.CreateCell(row.Cells.Count);
            newCell.SetCellValue(cellValue);
            newCell.CellStyle = cellstyle;
            return newCell;
        }
        private static Excel.ICell AddCellToRow(Excel.IRow row, string cellValue, Excel.ICellStyle cellstyle)
        {
            Excel.ICell newCell = row.CreateCell(row.Cells.Count);
            newCell.SetCellValue(cellValue);
            newCell.CellStyle = cellstyle;
            return newCell;
        }
        private static Excel.ICell AddCellToRow(Excel.IRow row, DateTime cellValue, Excel.ICellStyle cellstyle)
        {
            Excel.ICell newCell = row.CreateCell(row.Cells.Count);
            newCell.SetCellValue(cellValue);

            newCell.SetCellType(Excel.CellType.String);
            newCell.CellStyle = cellstyle;

            return newCell;
        }

        #endregion
#endif
        string CreateHTMLTableForWorkChangeRequest(Dictionary<WorkChangeRequest, string> dictWorkChangeRequestsProblem)
        {


            StringBuilder strBld = new StringBuilder();

            strBld.Append(string.Format(@"<!--Container Table-->
    <table  cellpadding = ""0"" cellspacing = ""0"" border = ""0"" width = ""99 % "">
             <tr>
               <!--Container Table-->
           <table  cellpadding = ""5"" cellspacing = ""0"" border = ""1"" width = ""10 % "">
                    <tr><td colspan = ""5"" bgcolor ={1} align=""center"">Расцветка фона по состоянию запроса</td></tr>
  <tr>
    <td bgcolor ={3} style=align=""center"" nowrap> &nbsp;Новый запрос&nbsp; </td>
    <td bgcolor ={2} style=align=""center"" nowrap> &nbsp;Принят РП&nbsp; </td>
    <td bgcolor ={5} style=align=""center"" nowrap> &nbsp;Отклонён РП&nbsp; </td>
    <td bgcolor ={1} style=align=""center"" nowrap> &nbsp;Запрос закрыт </td>
    <td bgcolor ={0} style=align=""center"" nowrap> &nbsp;Некорректные данные&nbsp; </td>
  </tr>
           </table>

           <p></p>
           <!--End Container Table -->
             </tr>
                 <tr>
                     <td align = ""left""></td>
                        <!--End Email Wrapper Table-->
           <table  cellpadding = ""0"" cellspacing = ""0"" border = ""1"" width = ""99 % "">
              <tr>
                <th nowrap bgcolor ={1}> № </th>
                <th nowrap bgcolor ={1}> Дата создания </th>
				<th nowrap bgcolor ={1}> Тип запроса </th>
				<th nowrap bgcolor ={1}> Содержание запроса </th>
                <th nowrap bgcolor ={1}> Название работы </th>
				<th nowrap bgcolor ={1}> Номер проекта </th>
				<th nowrap bgcolor ={1}> ФИО руководителя проекта </th>
				<th nowrap bgcolor ={1}> План подразделения </th>
				<th nowrap bgcolor ={1}> Решение РП </th>
				<th bgcolor ={1}> Ошибка </th>
				
			 </tr> ", BackGroundColor.LightBlue, BackGroundColor.LightGray, BackGroundColor.LightGreen, BackGroundColor.Pink, BackGroundColor.White, BackGroundColor.Yellow));

            int index = 1;
            foreach (var kvp in dictWorkChangeRequestsProblem)
            {
                WorkChangeRequest request = kvp.Key;
                string bgColor = BackGroundColor.LightBlue;
                if (request.StageName == WorkChangeRequest.Stages.New)
                { bgColor = BackGroundColor.Pink; }
                else if (request.StageName == WorkChangeRequest.Stages.DoneByManager)
                {
                    if (request.ManagerDecision == WorkChangeRequest.ManagerDecisionValue.Accept) bgColor = BackGroundColor.LightGreen;
                    else bgColor = BackGroundColor.Yellow;
                }
                else if (request.StageName == WorkChangeRequest.Stages.Completed)
                {
                    bgColor = BackGroundColor.LightGray;
                }
                string workName = "Не доступно";
                string workLink = "";
                ProjectManagementWork departmentWork = request.DepartmentWork;
                if (!departmentWork.IsEmpty)
                {
                    workName = departmentWork.Name;
                    workLink = HTML_MailMessage.GetLinkFor(departmentWork.ReferenceObject);
                }
                string link = HTML_MailMessage.GetLinkFor(request.ReferenceObject);
                strBld.Append(String.Format(@"
<tr bgcolor=""{0}"">
    <td>{1}</td>
    <td>{2}</td>
    <td>{3}</td>
    <td style=align=""right"">
        <a href={4}>
            {5}
        </a>
    </td>
<td style=align=""right"">
        <a href={6}>
            {7}
        </a>
    </td>
    <td>{8}</td>
    <td>{9}</td>
    <td>{10}</td>
    <td>{11}</td>
    <td>{12}</td>
    
</tr>",
                                            bgColor,//0
                                            index++,//1
                                            request.CreationDate.ToShortDateString(),//2
                                            request.Type,//3
                                            link, //4
                                            request.Content + "&nbsp",//5
                                            workLink,//6
                                            workName,//7
                                            request.ProjectNumberFrom1C + "&nbsp",//8
                                            request.ManagerName + "&nbsp",//9
                                            request.DepartmentProjectName + "&nbsp",//10
                                            request.ManagerDecision,//11
                                            kvp.Value + "&nbsp"//12
                                           ));
            }
            strBld.Append(@"</table>");

            return strBld.ToString();
        }

        private bool IsCurrentUserHaveRole(WorkChangeRequest workChangeRequest, string role)
        {
            bool result = false;
            User currentUser = References.Connection.ClientView.GetUser();// (User)CurrentUser;
            if (role == WorkChangeRequest.Role.Manager)
            {
                if (workChangeRequest.ManagerName == currentUser.FullName) result = true;
                else
                {
                    ProjectManagementWork departmentWork = workChangeRequest.DepartmentWork;

                    ProjectManagementWork managerWork = departmentWork.ProjectFrom1C.Project;
                    if (managerWork != null)
                    {
                        if (managerWork.Project.GetUsersHaveAccess(ProjectManagementWork.AccessCode.NoLimit).Contains(currentUser)) result = true;
                    }

#if test
                    else MessageBox.Show("managerWork = null");
#endif
                }
            }
            else if (role == WorkChangeRequest.Role.DepartmentScheduler)
            {
                if (workChangeRequest.DepartmentWork != null && !workChangeRequest.DepartmentWork.IsEmpty)
                {
                    if (workChangeRequest.DepartmentWork.Project.GetUsersHaveAccess(ProjectManagementWork.AccessCode.NoLimit).Contains(currentUser)) result = true;
                }
            }
            return result;
        }


        /// <summary>
        /// Метод проверки статуса и вызова необходимых действий при необходимости
        /// Запускается по кнопке или по расписанию
        /// </summary>
        public void ПроверитьСтатусВсехНезакрытыхЗапросов()
        {
            //получаем список всех не закрытых запросов
            List<WorkChangeRequest> listWorkChangeRequests = GetNotCompletedRequests();
            Dictionary<WorkChangeRequest, string> listWCRsWithMissedManagersWork = new Dictionary<WorkChangeRequest, string>();
            Dictionary<WorkChangeRequest, string> listRequestsWithMissedDepartmentWork = new Dictionary<WorkChangeRequest, string>();
            //Проверяем каждый запрос в зависимости от текущего статуса
            foreach (var workChangeRequest in listWorkChangeRequests)
            {
                try
                {
                    CheckRequestStatus(workChangeRequest);
                }

                catch (ManagerWorkNotFoundException e)
                { listWCRsWithMissedManagersWork.Add(workChangeRequest, e.ToString()); }
                catch (Exception e)
                { listRequestsWithMissedDepartmentWork.Add(workChangeRequest, e.ToString()); }
            }
            if (listRequestsWithMissedDepartmentWork.Count > 0)
                SendMessageToAdminAboutProblem(listRequestsWithMissedDepartmentWork);
            if (listWCRsWithMissedManagersWork.Count > 0)
                SendMessageToPlanningDepartmentsAboutProblem(listWCRsWithMissedManagersWork);

        }


        /// <summary>
        /// Метод проверки статуса и вызова необходимых действий при необходимости
        /// Запускается по кнопке 
        /// </summary>
        public void ПроверитьСтатусВыбранногоЗапроса()
        {
            ReferenceObject current = Context.ReferenceObject;
            WorkChangeRequest workChangeRequest = Factory.Create_WorkChangeRequest(current);
            try
            {
                CheckRequestStatus(workChangeRequest);
            }
            catch (WorkNotFoundException e)
            { MessageBox.Show(e.ToString()); }
        }

        /// <summary>
        /// проверяет текущий статус запроса, при необходимости меняет его и вызывает положенные действия
        /// </summary>
        /// <param name="workChangeRequest"></param>
        private void CheckRequestStatus(WorkChangeRequest workChangeRequest)
        {
            string changeTypeName = workChangeRequest.Type;

            bool emptyWork = workChangeRequest.DepartmentWork.IsEmpty;
            if (emptyWork)//если работы в плане подразделения нет, то закрываем
            {
                workChangeRequest.Close(WorkChangeRequest.CompleteNote.AutoClose);
            }

            //если работа подразделения перенесена в архив то закрываем
            //если работа РП перенесена в архив то закрываем
            if (Is_DepartmentWorkOrManagerWork_Archieved(workChangeRequest))
            {
                workChangeRequest.CompleteWithNewEmptyEditReasonNote(WorkChangeRequest.CompleteNote.WorkArchived);
                // workChangeRequest.Close(WorkChangeRequest.CompleteNote.WorkArchived);
            }

            else
            {
                //если последняя запись в ЖР - пустая, то можно закрывать
                EditReasonNote lastEditReasonNote = GetLastEditReasonNote(workChangeRequest.DepartmentWork);
                bool isLastEditReasonNoteEmpty = !lastEditReasonNote.IsActionOfManagerNeeded || (lastEditReasonNote.TypeCode == EditReasonNote.Type.NotDefined && string.IsNullOrWhiteSpace(lastEditReasonNote.Reason));
                if (workChangeRequest.CanBeClosed || isLastEditReasonNoteEmpty)
                {
                    workChangeRequest.Close(WorkChangeRequest.CompleteNote.AutoClose);
                }

                else if (workChangeRequest.StageName == WorkChangeRequest.Stages.New)
                {

                    if (workChangeRequest.IsChangeStageCanBeComplete)//проверяем выполнил ли действия РП
                    {
                        workChangeRequest.CompleteByManager(WorkChangeRequest.CompleteNote.AutoComplete);
                    }
                }
            }
            //else if (workChangeRequest.StageName == WorkChangeRequest.Stages.DoneByManager && //стадия обработано РП
            //         workChangeRequest.ManagerDecision == WorkChangeRequest.ManagerDecisionValue.Accept && // Решение РП = Принять
            //         workChangeRequest.CanChangeStageForAcceptedRequestToClose)//проверяем выполнил ли действия ПлОП
            //{
            //    workChangeRequest.Close(WorkChangeRequest.CompleteNote.AutoComplete);
            //}
            //else throw new Exception("Не применимая стадия.");
        }

        /// <summary>
        /// Находится ли связанная работа или синхронизированная работа в архиве
        /// </summary>
        /// <param name="workChangeRequest"></param>
        /// <returns></returns>
        private bool Is_DepartmentWorkOrManagerWork_Archieved(WorkChangeRequest workChangeRequest)
        {
            bool isArchieved = false;
            isArchieved = Is_Work_Archieved(workChangeRequest.DepartmentWork);
            if (!isArchieved)
            {
                ProjectManagementWork managerWork = workChangeRequest.GetManagerWork();
                if (managerWork != null) isArchieved = Is_Work_Archieved(managerWork);
            }
            return isArchieved;
        }

        /// <summary>
        /// Находится ли работа в архиве
        /// </summary>
        /// <param name="work"></param>
        /// <returns></returns>
        private bool Is_Work_Archieved(ProjectManagementWork work)
        {
            bool isArchieved = false;
            if (!work.IsEmpty) isArchieved = IsInFolderNamed(work, "Дополнительно");
            return isArchieved;
        }

        private bool IsInFolderNamed(ProjectTreeItem work, string folderName)
        {
            bool result = false;
            ProjectFolder projectFolder = work as ProjectFolder;
            if (projectFolder != null && projectFolder.Name == folderName) return true;
            if (work.Parent != null) result = IsInFolderNamed(work.Parent as ProjectTreeItem, folderName);
            return result;
        }

        /// <summary>
        /// Возвращает последнюю запись из Журнала Редактирования для работы
        /// если записей нет то null
        /// </summary>
        /// <param name="departmentWork"></param>
        /// <returns></returns>
        private EditReasonNote GetLastEditReasonNote(ProjectManagementWork departmentWork)
        {
            ParameterGroup linkToWork = EditReasonNote.Reference.ParameterGroup.ReferenceInfo.Description.OneToOneLinks.First(link => link.Guid == EditReasonNote.ERN_link_ToPrjctMgmntN1_Guid);

            Filter filter = new Filter(EditReasonNote.Reference.ParameterGroup.ReferenceInfo);

            // Условие: гуид связанной работы равен departmentWork.Guid
            ReferenceObjectTerm term2 = new ReferenceObjectTerm(filter.Terms);
            // добавляем связь в путь к параметру
            term2.Path.AddGroup(linkToWork);
            term2.Path.AddParameter(linkToWork.SlaveGroup[SystemParameterType.Guid]);
            // список операторов сравнения, которые поддерживает прарамeтр можно получить у него - ParamterInfo.GetComparisonOperators()
            term2.Operator = ComparisonOperator.Equal;
            term2.Value = departmentWork.Guid;

            List<ReferenceObject> listEditReasonNotes = EditReasonNote.Reference.Find(filter);
            if (listEditReasonNotes.Count == 0) return null;
            DateTime lastDateTime = DateTime.MinValue;
            ReferenceObject tempReferenceObject = null;
            foreach (var item in listEditReasonNotes)
            {
                if (item.SystemFields.CreationDate > lastDateTime)
                {
                    tempReferenceObject = item;
                    lastDateTime = item.SystemFields.CreationDate;
                }
            }
            EditReasonNote lastEditReasonNote = new EditReasonNote(tempReferenceObject);
            return lastEditReasonNote;
        }

        /// <summary>
        /// Возвращает список не закрытых запросов (все кроме стадии "Закрыт")
        /// </summary>
        /// <returns></returns>
        private List<WorkChangeRequest> GetNotCompletedRequests()
        {
            //находим стадию "Завершено"
            Stage stageCompleted = Stage.GetStages(References.WorkChangeRequestReferenceInfo).Where(s => s.Name == WorkChangeRequest.Stages.Completed).FirstOrDefault();

            //Создаем фильтр
            Filter filter = new Filter(References.WorkChangeRequestReferenceInfo);
            //Добавляем условие поиска – «Стадия != Закрыт»
            filter.Terms.AddTerm(References.WorkChangeRequestReference.ParameterGroup[SystemParameterType.Stage],
                                 ComparisonOperator.NotEqual, stageCompleted);

            //Получаем список объектов, в качестве условия поиска – сформированный фильтр
            List<ReferenceObject> listObj = References.WorkChangeRequestReference.Find(filter);
            List<WorkChangeRequest> result = new List<WorkChangeRequest>();
            foreach (var item in listObj)
            {
                result.Add(Factory.Create_WorkChangeRequest(item));
            }
            return result;
        }




        /// <summary>
        /// Рассылка о закрытии запроса
        /// </summary>
        /// <param name="workChangeRequest"></param>
        private void SendMessageAboutClose(WorkChangeRequest workChangeRequest)
        {
            List<ReferenceObject> mailRecipients = new List<ReferenceObject>();
            References.ProjectManagementReference.LoadSettings.LoadDeleted = true;
            References.ProjectManagementReference.Refresh();

            try
            {
                ProjectManagementWork departmentWork = workChangeRequest.DepartmentWork;

                if (!departmentWork.IsEmpty)
                {
                    ProjectManagementWork projectFrom1C = departmentWork.ProjectFrom1C;
                    if (projectFrom1C != null && !projectFrom1C.IsEmpty)
                    {
                        ProjectManagementWork managerProject = projectFrom1C.Project;
                        List<ProjectManagementWork> listMasterWorks = Synchronization.GetSynchronizedWorksFromSpace(departmentWork, PlanningSpaceGuidString.WorkingPlans, true);
                        if (!managerProject.IsEmpty)
                        {
                            mailRecipients.AddRange(managerProject.GetUsersHaveAccess(ProjectManagementWork.AccessCode.NoLimit));//РП (все кто имеет доступ без ограничений к плану РП)
                        }
                    }
                    EditReasonNote initialEditReasonNote = new EditReasonNote(workChangeRequest.EditReasonUID);
                    User initiator = initialEditReasonNote.Author;//автор записи в журнале редактирования
                    mailRecipients.Add(initiator);
                    mailRecipients.AddRange(departmentWork.Project.GetUsersHaveAccess(ProjectManagementWork.AccessCode.NoLimit));//ПлОП (все кто имеет доступ без ограничений к плану подразделения)

#if test
                    mailRecipients = null;
                    //mailRecipients = new List<ReferenceObject>{ Lipaev };
                    //mailRecipients = new List<ReferenceObject> { (User)CurrentUser };
#endif
                    if (mailRecipients != null && mailRecipients.Count > 0)
                    {
                        string text = "Запрос закрыт <br>" + CreateHTMLTableForWorkChangeRequest(new List<WorkChangeRequest> { workChangeRequest });
                        string subject = "Запрос закрыт";
                        HTML_MailMessage message = new HTML_MailMessage(subject, text, mailRecipients, new List<ReferenceObject> { workChangeRequest.ReferenceObject });
                        message.Send();
                    }
                }
#if test && !server
                else
                {
                    MessageBox.Show("departmentWork.IsEmpty");
                }
#endif
            }
            finally
            {
                References.ProjectManagementReference.LoadSettings.LoadDeleted = false;
                References.ProjectManagementReference.Refresh();
            }
        }

        struct BackGroundColor
        {
            public const string LightBlue = "#62ebff";
            public const string LightGray = "#e1e1e1";
            public const string LightGreen = "#0ffda1";
            public const string Pink = "#ff859c";
            public const string White = "#ffffff";
            public const string Yellow = "#fff962";
        }

        string CreateHTMLTableForWorkChangeRequest(List<WorkChangeRequest> listRequests)
        {
            StringBuilder strBld = new StringBuilder();

            strBld.Append(string.Format(@"<!--Container Table-->
    <table  cellpadding = ""0"" cellspacing = ""0"" border = ""0"" width = ""99 % "">
             <tr>
               <!--Container Table-->
           <table  cellpadding = ""5"" cellspacing = ""0"" border = ""1"" width = ""10 % "">
                    <tr><td colspan = ""5"" bgcolor ={1} align=""center"">Расцветка фона по состоянию запроса</td></tr>
  <tr>
    <td bgcolor ={3} style=align=""center"" nowrap> &nbsp;Новый запрос&nbsp; </td>
    <td bgcolor ={2} style=align=""center"" nowrap> &nbsp;Принят РП&nbsp; </td>
    <td bgcolor ={5} style=align=""center"" nowrap> &nbsp;Отклонён РП&nbsp; </td>
    <td bgcolor ={1} style=align=""center"" nowrap> &nbsp;Запрос закрыт </td>
    <td bgcolor ={0} style=align=""center"" nowrap> &nbsp;Некорректные данные&nbsp; </td>
  </tr>
           </table>

           <p></p>
           <!--End Container Table -->
             </tr>
                 <tr>
                     <td align = ""left""></td>
                        <!--End Email Wrapper Table-->
           <table  cellpadding = ""0"" cellspacing = ""0"" border = ""1"" width = ""99 % "">
              <tr>
                <th nowrap bgcolor ={1}> № </th>
                <th nowrap bgcolor ={1}> Дата создания </th>
				<th nowrap bgcolor ={1}> Тип запроса </th>
				<th nowrap bgcolor ={1}> Содержание запроса </th>
                <th nowrap bgcolor ={1}> Название работы (Исполнитель) </th>
				<th nowrap bgcolor ={1}> Номер проекта </th>
				<th nowrap bgcolor ={1}> ФИО руководителя проекта </th>
				<th nowrap bgcolor ={1}> План подразделения </th>
				<th nowrap bgcolor ={1}> Решение РП </th>
				<th bgcolor ={1}> Причина отклонения </th>
				<th bgcolor ={1}> Комментарий </th>
			 </tr> ", BackGroundColor.LightBlue, BackGroundColor.LightGray, BackGroundColor.LightGreen, BackGroundColor.Pink, BackGroundColor.White, BackGroundColor.Yellow));

            int index = 1;
            foreach (var request in listRequests)
            {
                string bgColor = BackGroundColor.LightBlue;
                if (request.StageName == WorkChangeRequest.Stages.New)
                { bgColor = BackGroundColor.Pink; }
                else if (request.StageName == WorkChangeRequest.Stages.DoneByManager)
                {
                    if (request.ManagerDecision == WorkChangeRequest.ManagerDecisionValue.Accept) bgColor = BackGroundColor.LightGreen;
                    else bgColor = BackGroundColor.Yellow;
                }
                else if (request.StageName == WorkChangeRequest.Stages.Completed)
                {
                    bgColor = BackGroundColor.LightGray;
                }

                string workName = "Не доступно";
                string workLink = "";
                string workUsedResources = "Не указан";
                ProjectManagementWork departmentWork = request.DepartmentWork;
                if (!departmentWork.IsEmpty)
                {
                    workName = departmentWork.Name;
                    workLink = HTML_MailMessage.GetLinkFor(departmentWork.ReferenceObject);
                    workUsedResources = departmentWork.UsedResourcesNames;
                }

                string link = HTML_MailMessage.GetLinkFor(request.ReferenceObject);
                strBld.Append(String.Format(@"
<tr bgcolor=""{0}"">
    <td>{1}</td>
    <td>{2}</td>
    <td>{3}</td>
    <td style=align=""right"">
        <a href={4}>
            {5}
        </a>
    </td>
    <td style=align=""right"">
        <a href={6}>
            {7}
        </a> (Исполнитель: {8})
    </td>
    <td>{9}</td>
    <td>{10}</td>
    <td>{11}</td>
    <td>{12}</td>
    <td>{13}</td>
    <td>{14}</td>
</tr>",
                                            bgColor,//0
                                            index++,//1
                                            request.CreationDate.ToShortDateString(),//2
                                            request.Type,//3
                                            link, //4
                                            request.Content + "&nbsp",//5
                                            workLink,//6
                                            workName,//7
                                            "&nbsp" + workUsedResources,//8
                                            request.ProjectNumberFrom1C + "&nbsp",//9
                                            request.ManagerName + "&nbsp",//10
                                            request.DepartmentProjectName + "&nbsp",//11
                                            request.ManagerDecision + "&nbsp",//12
                                            request.ReasonToReject + "&nbsp",//13
                                            request.CommentToComplete + "&nbsp"//14
                                           ));
            }
            strBld.Append(@"</table>");

            return strBld.ToString();
        }

        /// <summary>
        /// Рассылка о новом запросе
        /// </summary>
        /// <param name="workChangeRequest"></param>
        private void SendMessageAboutNewRequest(WorkChangeRequest workChangeRequest)
        {
            List<ReferenceObject> mailRecipients = new List<ReferenceObject>();

            ProjectManagementWork departmentWork = workChangeRequest.DepartmentWork;
            if (!departmentWork.IsEmpty)
            {

                List<ProjectManagementWork> listMasterWorks = Synchronization.GetSynchronizedWorksFromSpace(departmentWork, PlanningSpaceGuidString.WorkingPlans, true);
                if (listMasterWorks.Count > 0)
                {
                    //if (listMasterWorks.Count > 1) { throw new InvalidOperationException("Укрупнений больше одного!"); }

                    ProjectManagementWork managerWork = listMasterWorks.First();

                    mailRecipients.AddRange(managerWork.Project.GetUsersHaveAccess(ProjectManagementWork.AccessCode.NoLimit)); //добавляем РП и помощников
                }

#if test
                mailRecipients = null;
                //mailRecipients = new List<ReferenceObject>{ Lipaev };
                //mailRecipients = new List<ReferenceObject> { (User)CurrentUser };
#endif
                if (mailRecipients != null && mailRecipients.Count > 0)
                {
                    string text = "Новый запрос на изменение проекта <br>" + CreateHTMLTableForWorkChangeRequest(new List<WorkChangeRequest> { workChangeRequest });
                    string subject = "Новый запрос на изменение проекта";
                    HTML_MailMessage message = new HTML_MailMessage(subject, text, mailRecipients, new List<ReferenceObject> { workChangeRequest.ReferenceObject });
                    message.Send();

                }
            }

        }

        /// <summary>
        /// Рассылка о необходимости ручного закрытия запроса
        /// </summary>
        /// <param name="workChangeRequest"></param>
        private void SendMessageNeedManualClose(WorkChangeRequest workChangeRequest)
        {
            List<ReferenceObject> mailRecipients = new List<ReferenceObject>();

            ProjectManagementWork departmentWork = workChangeRequest.DepartmentWork;
            mailRecipients.AddRange(departmentWork.Project.GetUsersHaveAccess(ProjectManagementWork.AccessCode.NoLimit));
#if test
            mailRecipients = new List<ReferenceObject>() { Lipaev };
            //mailRecipients = new List<ReferenceObject> { (User)CurrentUser };
#endif
            string text = "Требуются ваши действия по закрытию запроса <br>" + CreateHTMLTableForWorkChangeRequest(new List<WorkChangeRequest> { workChangeRequest });
            string subject = "Требуются ваши действия по закрытию запроса";
            HTML_MailMessage message = new HTML_MailMessage(subject, text, mailRecipients, new List<ReferenceObject> { workChangeRequest.ReferenceObject });
            message.Send();
        }

        /// <summary>
        /// Возвращает значение true если необходимо создать запрос по этой записи
        /// </summary>
        /// <param name="editReasonNote"></param>
        /// <returns>true если необходимо создать запрос, в противном случае false</returns>
        private static bool IsNeedToCreateChangeRequest(EditReasonNote editReasonNote)
        {
            bool needToCreateChangeRequest = false;
            if (editReasonNote.IsActionOfManagerNeeded && //проверяем флаг "Требуется ответ РП"
                WorkChangeRequest.EnumTypes.Contains(editReasonNote.TypeName))//проверяем тип редактирования
            {
                needToCreateChangeRequest = true;
                //if (editReasonNote.TypeCode == EditReasonNote.Type.AddChildWork)
                //{
                //    needToCreateChangeRequest = true;
                //}
                //else
                //{ 
                //    //проверяем наличие синхронизированной работы РП
                //    List<ProjectManagementWork> listSyncWorks = Synchronization.GetSyncronizedWorksFromSpace(editReasonNote.Work, PlanningSpaceGuidString.WorkingPlans, true);
                //    if (listSyncWorks.Count > 0) needToCreateChangeRequest = true;

                //}
            }

            return needToCreateChangeRequest;
        }

        /// <summary>
        /// Проверяет находится ли работа в плане подразделений
        /// </summary>
        /// <param name="work"></param>
        /// <returns></returns>
        private static bool WorkNotFromDepartmentPlan(ProjectManagementWork work)
        {
            return work.PlanningSpace.ToString() != PlanningSpaceGuidString.Departments;
        }

        /// <summary>
        /// Закрывает все незакрытые запросы коррекции для работы
        /// </summary>
        /// <param name="work"></param>
        private void CloseAllUncompletedChangeRequestsForWork(ProjectManagementWork work)
        {
            List<WorkChangeRequest> listChangeRequest = GetUncompletedChangeRequestsForWork(work);
            foreach (var cr in listChangeRequest)
            {
                //if (cr.CanChangeStageForAcceptedRequestToClose)
                {
                    cr.Close("Закрыто в подразделении.");
                }
            }
        }

        /// <summary>
        /// Находит все незакрытые запросы коррекции для работы work
        /// </summary>
        /// <param name="work"></param>
        /// <returns>Список незакрытых запросов</returns>
        private static List<WorkChangeRequest> GetUncompletedChangeRequestsForWork(ProjectManagementWork work)
        {


            //MessageBox.Show(filter.ToString());
            List<ReferenceObject> listObj = GetWorkCRsReferenceObjects(work, null, new List<string> { WorkChangeRequest.Stages.New, WorkChangeRequest.Stages.DoneByManager });
            //MessageBox.Show("listObj.Count = " + listObj.Count.ToString());
            List<WorkChangeRequest> listChangeRequest = new List<WorkChangeRequest>();
            foreach (var item in listObj)
            {
                WorkChangeRequest wcr = Factory.Create_WorkChangeRequest(item);
                listChangeRequest.Add(wcr);
            }

            return listChangeRequest;
        }

        /// <summary>
        /// Находит все запросы коррекции для работы work
        /// </summary>
        /// <param name="work"></param>
        /// <returns>Список всех запросов</returns>
        private static List<WorkChangeRequest> GetChangeRequests(ProjectManagementWork work, EditReasonNote editReasonNote, IEnumerable<string> stageNames)
        {
            List<ReferenceObject> listObj = GetWorkCRsReferenceObjects(work, editReasonNote, stageNames);
            //MessageBox.Show("listObj.Count = " + listObj.Count.ToString());
            List<WorkChangeRequest> listChangeRequest = new List<WorkChangeRequest>();
            foreach (var item in listObj)
            {
                WorkChangeRequest wcr = Factory.Create_WorkChangeRequest(item);
                listChangeRequest.Add(wcr);
            }

            return listChangeRequest;
        }

        private static List<ReferenceObject> GetWorkCRsReferenceObjects(ProjectManagementWork work, EditReasonNote editReasonNote, IEnumerable<string> stageNames)
        {
            //Создаем ссылку на справочник
            Reference reference = References.WorkChangeRequestReference;
            ReferenceInfo info = References.WorkChangeRequestReferenceInfo;


            //Создаем фильтр
            Filter filter = new Filter(info);
            if (stageNames != null && stageNames.Count() > 0)
            {
                List<Stage> stages = new List<Stage>();
                foreach (var stageName in stageNames)
                {
                    //Находим стадию по имени
                    Stage stage = Stage.GetStages(info).Where(s => s.Name == stageName).FirstOrDefault();
                    stages.Add(stage);
                }
                //Добавляем условие поиска – «стадия входит в список stages»
                filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Stage],
                                     ComparisonOperator.IsOneOf, stages);
            }
            if (work != null && !work.IsEmpty)
            {
                //Добавляем условие поиска – «Guid связанной работы == Guid work»
                filter.Terms.AddTerm(reference.ParameterGroup[WorkChangeRequest.param_PMWorkUID_Guid],
                                     ComparisonOperator.Equal, work.Guid);
            }
            if (editReasonNote != null)
            {
                //Добавляем условие поиска – «Guid editReasonNote == Guid editReasonNote»
                filter.Terms.AddTerm(reference.ParameterGroup[WorkChangeRequest.param_EditReasonUID_Guid],
                                     ComparisonOperator.Equal, editReasonNote.Guid);
            }
            //Получаем список объектов, в качестве условия поиска – сформированный фильтр
            //MessageBox.Show(filter.ToString());
            List<ReferenceObject> listObj = reference.Find(filter);
            return listObj;
        }

        public bool CopyChangeReason(ReferenceObject work, ReferenceObject cr)
        {
            //MessageBox.Show(work.ToString() + "\n" + cr.ToString());
            ReferenceObject newCR = cr.CreateCopy();
            //newCR.BeginChanges();
            newCR.SetLinkedObject(EditReasonNote.ERN_link_ToPrjctMgmntN1_Guid, work);
            newCR.EndChanges();
            return true;
        }

        User Lipaev { get { return References.UsersReference.FindUser("Липаев Алексей Александрович"); } }
    }

    /// <summary>
    /// Для поиска записей лога изменения работы
    /// ver 1.0
    /// </summary>
    class WorkLogSearch
    {
        public static LogJournal Search(ProjectManagementWork work)
        {
            return Search(work.Guid);
        }

        public static LogJournal Search(Guid workGuid)
        {
            return new LogJournal(workGuid);
        }
    }

    /// <summary>
    /// Журнал изменений работы
    /// ver 2.0
    /// </summary>
    class LogJournal
    {
        public struct ParameterName
        {
            public const string EndDate = "Окончание";
            public const string StartDate = "Начало";
        }
        private Guid workGuid;
        List<ReferenceObject> ReferenceObjects;
        bool DataNotFound
        {
            get
            {
                if (this.ReferenceObjects == null || this.ReferenceObjects.Count == 0) return true;
                return false;
            }
        }

        public LogJournal(Guid workGuid)
        {
            this.workGuid = workGuid;

            ParameterInfo pi = References.WorkLogReferenceInfo.Description.OneToOneParameters.Find(new Guid("cc39ac2e-2278-4115-a140-06250d1713dc"));
            References.WorkLogReference.Refresh();
            this.ReferenceObjects = References.WorkLogReference.Find(pi, workGuid.ToString());
        }

        List<LogRecord> _AllRecords;
        /// <summary>
        /// Список всех изменений
        /// </summary>
        public IEnumerable<LogRecord> AllRecords
        {
            get
            {
                if (_AllRecords == null)
                {
                    _AllRecords = new List<LogRecord>();

                    if (!DataNotFound)
                    {
                        foreach (var item in this.ReferenceObjects)
                        {
                            _AllRecords.Add(new LogRecord(item));
                        }
                    }
                }
                return _AllRecords;
            }
        }

        /// <summary>
        /// Возвращает дату окончания действовавшую на дату запроса, если данных нет то null
        /// </summary>
        /// <param name="creationDate"></param>
        /// <returns></returns>
        internal DateTime? GetEndDateValue(DateTime creationDate)
        {
            DateTime? result = null;
            List<LogRecord> sortedList = this.AllRecords.Where(l => l.Content.ToLower().Contains(ParameterName.EndDate.ToLower())).ToList();
            sortedList.Sort(delegate (LogRecord x, LogRecord y)//сортируем по убыванию даты записи
            {
                return y.CreationDate.CompareTo(x.CreationDate);
            });
            //sortedList = sortedList.Where(l => l.Content.ToLower().Contains("дата окончания")).ToList();
            string text = "";
            DateTime tempEndDate = DateTime.MinValue;
            foreach (LogRecord logRecord in sortedList)
            {
                //находим первую запись с датой создания меньше заданной и берём оттуда новое значение
                if (logRecord.CreationDate < creationDate)
                {

                    result = DateTime.Parse(logRecord.NewValue);
                    break;
                }
                //MessageBox.Show(logRecord.Content + " - Content\n" + logRecord.CreationDate.ToString() + " -  CreationDate\n" + logRecord.OldValue);
            }
            if (result == null && sortedList.Count > 0)
            {
                //берём последнюю, самую старую, запись и берём из неё значение "было"
                string value = sortedList.Last().OldValue;
                if (!string.IsNullOrWhiteSpace(value))
                    result = DateTime.Parse(value);
            }
            return result;
        }
        //private static readonly Guid param_LinkedObjectGuid_Guid = new Guid("14068df7-a895-4aff-9698-95527abca35b");     //Guid параметра "Guid объекта" Строка переменной длины(Длина 60 символов) GuidObj
        //private static readonly Guid list_Records_Guid = new Guid("6442ddd9-a890-4d96-8231-eccb0fa188ce");     //Guid списка "Записи"
    }

    class LogRecord
    {
        public LogRecord(ReferenceObject referenceObject)
        {
            this.ReferenceObject = referenceObject;
        }

        public string OldValue
        {
            get
            {
                return this.ReferenceObject[param_OldValue_Guid].Value.ToString();
            }
        }

        /// <summary>
        /// Описание изменения
        /// </summary>
        public string Content
        {
            get
            {
                return this.ReferenceObject[param_Content_Guid].Value.ToString();
            }
        }

        public string NewValue
        {
            get
            {
                return this.ReferenceObject[param_NewValue_Guid].Value.ToString();
            }
        }

        /// <summary>
        /// Дата создания записи
        /// </summary>
        public DateTime CreationDate
        { get { return this.ReferenceObject.SystemFields.CreationDate; } }

        public override string ToString()
        {
            return string.Format("Дата создания - {0}; описание - {1}; было -{2}; стало - {3}",
                this.CreationDate.ToString(), this.Content, this.OldValue, this.NewValue);
        }

        ReferenceObject ReferenceObject;

        private static readonly Guid param_OldValue_Guid = new Guid("7a0c3e74-4f48-4416-9016-64de9e28b136");     //Guid параметра "Было" HTML-текст
        private static readonly Guid param_NewValue_Guid = new Guid("ef99e76d-7fbb-4fe6-a01e-fb19415bec00");     //Guid параметра "Стало" HTML-текст
        private static readonly Guid param_Content_Guid = new Guid("b7d54497-993c-4f54-a40e-54d963b3e4c1");     //Guid параметра "Описание события" Строка переменной длины	

    }

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

    internal abstract class WorkChangeRequest
    {
        public WorkChangeRequest(EditReasonNote editReasonNote)
        {
            string type = editReasonNote.TypeName ?? "null";
            //проверяем допустимость типа редактирования
            if (!EnumTypes.Contains(type)) throw new Exception("Недопустимый тип редактирования - " + type);
            if (!editReasonNote.IsActionOfManagerNeeded) throw new Exception(string.Format("Не установлен флаг \"Требуются действия РП\", Reference - {0}, ObjectGuid - {1}", EditReasonNote.Reference.Name, editReasonNote.Guid.ToString()));
            //заполняем его данными из журнала редактирования
            this.CommentToComplete = "";
            this.Content = editReasonNote.Reason; ;
            this.DepartmentProjectName = editReasonNote.Work.ProjectName;
            this.EditReasonUID = editReasonNote.Guid;
            this.ManagerName = editReasonNote.Work.ProjectResponsibleNameFrom1C;
            this.Name = editReasonNote.Work.Name;
            this.DepartmentWorkGUID = editReasonNote.Work.Guid;
            this.ProjectNumberFrom1C = editReasonNote.Work.ProjectNumberFrom1C;
            this.ManagerDecision = "Не обработан";
            this.ReasonToReject = "";
            this.Type = type;
            this.CreationDate = editReasonNote.CreateDate;
            this.SaveToDB();

        }
        public WorkChangeRequest(ReferenceObject workChangeRequest_ro)
        {
            this.ReferenceObject = workChangeRequest_ro;
            this._CompletedByManagerDate = workChangeRequest_ro[param_CompletedByManager_Guid].GetDateTime();
            this._CloseDate = workChangeRequest_ro[param_CompleteDate_Guid].GetDateTime();
            this.EditReasonUID = workChangeRequest_ro[param_EditReasonUID_Guid].GetGuid();
            this.Type = workChangeRequest_ro[param_RequestType_Guid].GetString();
            this.ReasonToReject = workChangeRequest_ro[param_ReasonToReject_Guid].GetString();
            this.Content = workChangeRequest_ro[param_ChangeRequestContent_Guid].GetString();
            this.DepartmentWorkGUID = workChangeRequest_ro[param_PMWorkUID_Guid].GetGuid();
            this.ProjectNumberFrom1C = ReferenceObject[param_ProjectNumber_Guid].GetString();
            this.CommentToComplete = workChangeRequest_ro[param_CommentToComplete_Guid].GetString();
            this.ManagerName = workChangeRequest_ro[param_ManagerName_Guid].GetString();
            this.Name = workChangeRequest_ro[param_Name_Guid].GetString();
            this.DepartmentProjectName = workChangeRequest_ro[param_DepartmentProjectName_Guid].GetString();
            this.ManagerDecision = workChangeRequest_ro[param_ManagerDecision_Guid].GetString();
            this.CreationDate = workChangeRequest_ro[param_CreationDate_Guid].GetDateTime();
        }

        public ReferenceObject ReferenceObject
        {
            get; private set;
        }

        DateTime? _CreateDate;
        /// <summary>
        /// Дата создания запроса
        /// </summary>
        public DateTime CreationDate
        {
            get; set;
        }

        /// <summary>
        /// Комментарий закрытия запроса
        /// </summary>
        public string CommentToComplete { get; set; }

        DateTime? _CompletedByManagerDate;
        /// <summary>
        /// Дата обработки запроса Руководителем проекта
        /// </summary>
        public DateTime? CompletedByManagerDate
        {
            get
            {
                if (_CompletedByManagerDate == DateTime.MinValue) return null;
                else return _CompletedByManagerDate;
            }
            set
            {
                if (_CompletedByManagerDate != value)
                    _CompletedByManagerDate = value;
            }
        }

        DateTime? _CloseDate;
        /// <summary>
        /// Дата закрытия запроса
        /// </summary>
        public DateTime? CloseDate
        {
            get
            {
                if (_CloseDate == DateTime.MinValue) return null;
                else return _CloseDate;
            }
            //#if test
            set
            {
                if (_CloseDate != value)
                    _CloseDate = value;
            }
            //#endif
        }

        /// <summary>
        /// Содержание запроса
        /// </summary>
        public string Content { get; set; }

        /// <summary>
        /// Наиемнование  проекта связанной работы
        /// </summary>
        public string DepartmentProjectName { get; set; }

        public Guid EditReasonUID { get; set; }

        /// <summary>
        /// Решение РП
        /// </summary>
        public string ManagerDecision { get; set; }

        /// <summary>
        /// ФИО руководителя проекта
        /// </summary>
        public string ManagerName { get; set; }

        /// <summary>
        /// Номер проекта в 1С
        /// </summary>
        public string ProjectNumberFrom1C { get; set; }

        /// <summary>
        /// Наименование
        /// </summary>
        public string Name { get; set; }

        ProjectManagementWork _DepartmentWork;

        /// <summary>
        /// Связанная работа
        /// </summary>
        public ProjectManagementWork DepartmentWork
        {
            get
            {
                if (_DepartmentWork == null) _DepartmentWork = new ProjectManagementWork(this.DepartmentWorkGUID);
                return _DepartmentWork;
            }
        }

        /// <summary>
        /// Первая работа найденная по синхронизации, в пространстве "Рабочие планы",
        /// если не найдена то null
        /// </summary>
        public ProjectManagementWork GetManagerWork(bool includeDeleted = false)
        {

            if (this.DepartmentWork == null || this.DepartmentWork.IsEmpty) return null;
            if (includeDeleted)
                References.ProjectManagementReference.LoadSettings.LoadDeleted = true;
            ProjectManagementWork managerWork = Synchronization.GetSynchronizedWorksFromSpace(this.DepartmentWork, PlanningSpaceGuidString.WorkingPlans, true).FirstOrDefault();
            if (includeDeleted) References.ProjectManagementReference.LoadSettings.LoadDeleted = false;
            return managerWork;
        }
        public Guid DepartmentWorkGUID { get; private set; }

        /// <summary>
        /// Причина отклонения запроса
        /// </summary>
        public string ReasonToReject { get; set; }

        /// <summary>
        /// Стадия запроса
        /// </summary>
        public string StageName
        {
            get
            {
                string result = "";
                if (this.ReferenceObject != null) result = this.ReferenceObject.SystemFields.Stage.Stage.Name;
                return result;
            }
        }

        /// <summary>
        /// Тип запроса
        /// </summary>
        public string Type { get; set; }

        /// <summary>
        /// Может ли стадия быть изменена на "Закрыт"
        /// </summary>
        public abstract bool CanChangeStageForAcceptedRequestToClose { get; }

        /// <summary>
        /// Попытаться выпонить действия необходимые для закрытия запроса
        /// </summary>
        /// <returns></returns>
        public abstract bool TryAutomaticallyDoCloseActions();

        /// <summary>
        /// Может ли стадия быть изменена на "Обработан РП"
        /// </summary>
        public abstract bool IsChangeStageCanBeComplete { get; }

        public bool SaveToDB()
        {
            try
            {
                BeginChanges();
            }
            catch (Exception e) { throw new Exception(e.ToString() + string.Format("\n\r ID = {0}", this.ReferenceObject.SystemFields.Id)); }
            this.ReferenceObject[param_CompletedByManager_Guid].Value = CompletedByManagerDate;
            this.ReferenceObject[param_CreationDate_Guid].Value = this.CreationDate;
            this.ReferenceObject[param_CompleteDate_Guid].Value = CloseDate;
            this.ReferenceObject[param_RequestType_Guid].Value = Type;
            this.ReferenceObject[param_ReasonToReject_Guid].Value = ReasonToReject;
            this.ReferenceObject[param_ChangeRequestContent_Guid].Value = Content;
            this.ReferenceObject[param_PMWorkUID_Guid].Value = DepartmentWorkGUID;
            this.ReferenceObject[param_CommentToComplete_Guid].Value = CommentToComplete;
            this.ReferenceObject[param_ManagerName_Guid].Value = ManagerName;
            this.ReferenceObject[param_Name_Guid].Value = Name;
            this.ReferenceObject[param_DepartmentProjectName_Guid].Value = DepartmentProjectName;
            this.ReferenceObject[param_EditReasonUID_Guid].Value = EditReasonUID;
            this.ReferenceObject[param_ProjectNumber_Guid].Value = ProjectNumberFrom1C;
            this.ReferenceObject[param_ManagerDecision_Guid].Value = ManagerDecision;

            //MessageBox.Show("PMWorkUID" + this.PMWorkUID.ToString());
            //if (this.PMWork != null)
            //    this.ReferenceObject.SetLinkedObject(link_PMWork_Guid, this.PMWork.ReferenceObject);
            SaveChildParameters();
            return this.ReferenceObject.EndChanges();
        }

        protected virtual void SaveChildParameters()
        { }
        protected abstract void SaveChildOnCreatedParameters();

        public abstract ClassObject WorkChangeRequestClass { get; }
        public virtual bool CanBeClosed { get { return false; } }

        private void BeginChanges()
        {

            if (this.ReferenceObject == null)
            {
                ReferenceObject = References.WorkChangeRequestReference.CreateReferenceObject(WorkChangeRequestClass);
                SaveChildOnCreatedParameters();
            }
            else
            {
                this.ReferenceObject.Reload();
                if (this.ReferenceObject.LockState == ReferenceObjectLockState.LockedByCurrentUser)
                {
                    this.ReferenceObject.BeginChanges();
                }
                else if (this.ReferenceObject.LockState == ReferenceObjectLockState.LockedByOtherUser)
                {
                    this.ReferenceObject.Unlock();
                    this.ReferenceObject.BeginChanges();
                }
                else if (this.ReferenceObject.LockState == ReferenceObjectLockState.None)
                {
                    this.ReferenceObject.BeginChanges();
                }
            }

        }

        /// <summary>
        /// Набор действий достаточных для обработки РП
        /// </summary>
        internal abstract bool DoAcceptActions();


        internal void CompleteByManager(string comment)
        {
            //if (this.IsChangeStageCanBeComplete)
            //{
            //this.DoAcceptActions();
            this.ManagerDecision = ManagerDecisionValue.Accept;
            this._CompletedByManagerDate = DateTime.Now;
            this.ReasonToReject = comment;
            this.CommentToComplete = comment;
            this.SaveToDB();
            //this.ChangeStageTo(WorkChangeRequest.Stages.Completed);
            //}
            //else throw new Exception(string.Format("Не возможно изменить стадию запроса, не все условия выполнены. Reference - {0}, ObjectGuid - {1}", this.ReferenceObject.Reference.Name, this.ReferenceObject.SystemFields.Guid.ToString()));
        }

        /// <summary>
        /// Закрывает запрос созданием пустой записи в ЖРП связанной работы
        /// </summary>
        /// <param name="reason"></param>
        internal void CompleteWithNewEmptyEditReasonNote(string reason = "")
        {
            ProjectManagementWork departmentWork = this.DepartmentWork;//находим работу в плане подразделения
            EditReasonNote emptyEditReasonNote = new EditReasonNote(departmentWork, EditReasonNote.TypeNames[(int)EditReasonNote.Type.NotDefined], reason);
        }

        /// <summary>
        /// Набор действий достаточных для отклонения
        /// </summary>
        internal abstract void DoRejectActions();

        internal void RejectByManager(string reason, string comment)
        {
            this.DoRejectActions();
            this.ManagerDecision = ManagerDecisionValue.Reject;
            this._CompletedByManagerDate = DateTime.Now;
            this.ReasonToReject = reason;
            this.CommentToComplete = comment;
            this.SaveToDB();
            //this.ChangeStageTo(WorkChangeRequest.Stages.Completed);
        }

        /// <summary>
        /// Закрывает запрос с указанным комментарием
        /// </summary>
        /// <param name="comment"></param>
        internal void Close(string comment)
        {
            this._CloseDate = DateTime.Now;
            this.CommentToComplete = comment;
            this.SaveToDB();
        }

        public void ChangeStageTo(string stageName)
        {
            Stage newStage = Stage.GetStages(References.WorkChangeRequestReferenceInfo).Find(st => st.Name == stageName);
            newStage.Change(new List<ReferenceObject> { ReferenceObject });
            this.ReferenceObject.Reload();
        }

        public struct CompleteNote
        {
            public const string AutoClose = "Запрос закрыт автоматически";
            public const string AutoComplete = "Запрос обработан автоматически";
            public const string NewRequestCreation = "Запрос закрыт созданием нового запроса";
            public const string Cancel = "Запрос отклонён";
            public const string MissedWork = "Отсутствует связаная работа";
            public const string ManualClose = "Запрос закрыт вручную";
            public const string WorkArchived = "Работа, связанная с запросом, перенесена в архив";
        }

        public struct Role
        {
            public const string Manager = "РП";
            public const string DepartmentScheduler = "ПлОП";
        }

        public struct Stages
        {
            public const string New = "Новый";
            public const string DoneByManager = "Обработан РП";
            public const string Completed = "Закрыт";
        }

        public struct ManagerDecisionValue
        {
            public const string NotProcessed = "Не обработан";
            public const string Accept = "Принять";
            public const string Reject = "Отклонить";
        }


        /// <summary>
        /// Допустимые типы запроса
        /// </summary>
        public struct TypeName
        {
            public static readonly string DatesCorrection = EditReasonNote.TypeNames[(int)EditReasonNote.Type.DatesCorrection];
            public static readonly string ChangeWorkName = EditReasonNote.TypeNames[(int)EditReasonNote.Type.ChangeWorkName];
            public static readonly string ChangeResource = EditReasonNote.TypeNames[(int)EditReasonNote.Type.ChangeResource];
            public static readonly string DeleteChildWork = EditReasonNote.TypeNames[(int)EditReasonNote.Type.DeleteChildWork];
            public static readonly string AddChildWork = EditReasonNote.TypeNames[(int)EditReasonNote.Type.AddChildWork];
        }

        public static IEnumerable<string> EnumTypes = new List<string> {
            TypeName.DatesCorrection,
            TypeName.ChangeWorkName,
            TypeName.ChangeResource,
            TypeName.DeleteChildWork,
            TypeName.AddChildWork,
        };

        static readonly Guid param_CreationDate_Guid = new Guid("95b56f28-1dba-477e-bd14-896338e92da1"); //Guid параметра "Дата создания запроса" (может быть null)
        static readonly Guid param_CompleteDate_Guid = new Guid("33000878-bd5c-4c42-b5f8-f15091c49d02"); //Guid параметра "Дата закрытия" (может быть null)
        static readonly Guid param_CompletedByManager_Guid = new Guid("7d7dc9bc-6269-472f-9385-acefbed3a49a"); //Guid параметра "Обработано РП" (может быть null)
        static readonly Guid param_Name_Guid = new Guid("a6397b02-8704-4f24-80cd-00aacb03cd62"); //Guid параметра "Наименование" (Длина 255 символов)
        static readonly Guid param_DepartmentProjectName_Guid = new Guid("e2d56829-0d4e-4c15-a906-510edcb18a35"); //Guid параметра "Наименование плана подразделения" (Длина 255 символов)
        static readonly Guid param_ProjectNumber_Guid = new Guid("ad780748-4441-4173-84d0-aa9cbbde3283"); //Guid параметра "Номер проекта" (Длина 64 символов) Field_6002
        static readonly Guid param_ReasonToReject_Guid = new Guid("fb973537-20d9-4dcd-a7b2-cea5d230cbdc"); //Guid параметра "Причина отклонения запроса" (Длина 512 символов)    Field_6000
        static readonly Guid param_CommentToComplete_Guid = new Guid("9c752856-c6d0-4312-89ad-81800f129f6b"); //Guid параметра "Комментарий" Строка переменной длины(Длина 255 символов)    Field_6008
        static readonly Guid param_RequestType_Guid = new Guid("3027c606-32a3-4c2f-b7e5-2a2c5ebc6817"); //Guid параметра "Тип запроса" (Длина 64 символов) Field_6001
        static readonly Guid param_ManagerName_Guid = new Guid("c44bfd38-aeef-43c9-8106-8dc018ce4a1c"); //Guid параметра "ФИО Руководителя проекта" (Длина 255 символов)    Field_6003
        static readonly Guid param_ChangeRequestContent_Guid = new Guid("085fd94b-b855-410a-ae44-5362453dba36"); //Guid параметра "Содержание запроса" Текст переменной длины  Field_6005
        public static readonly Guid param_EditReasonUID_Guid = new Guid("8db18df0-a91a-495c-8ca9-c157d2a84ad3"); //Guid параметра "Идентификатор Записи в Журнале редактирования" Уникальный идентификатор Field_6009
        static readonly Guid param_ManagerDecision_Guid = new Guid("14001115-d6f4-49e2-8d08-ec314e92b595"); //Guid параметра "Решение РП" Строка переменной длины(Длина 25 символов) Field_6010

        public static readonly Guid param_PMWorkUID_Guid = new Guid("206337d3-aea3-421c-ae3c-18b29db73b3c"); //Guid параметра "Идентификатор работы" Уникальный идентификатор Field_6007
                                                                                                             //static readonly Guid link_PMWork_Guid = new Guid("cd00a92a-39c1-49ea-9b2c-d879528a870a"); //Guid Связь(список с одним) "Работа -> Управление проектами"  Link_2464
    }

    /// <summary>
    /// Запрос на изменение дат
    /// </summary>
    class WorkChangeRequest_DatesChange : WorkChangeRequest
    {
        public WorkChangeRequest_DatesChange(EditReasonNote note) : base(note)
        { }
        public WorkChangeRequest_DatesChange(ReferenceObject referenceObject) : base(referenceObject)
        { }

        public override bool CanChangeStageForAcceptedRequestToClose
        {
            get
            {
                bool result = this.StageName == Stages.DoneByManager && IsDepartmentManuallyDoneDatesCorrection();
                return result;
            }
        }

        private bool IsDepartmentManuallyDoneDatesCorrection()
        {
            //if (this.OriginalAndAimDatesCoincide) // если даты не были известны, то требуется ручная обработка ПлОП
            //    return false;//this.DepartmentWork.EndDate.Date> this.OriginalEndDate.Date && this.DepartmentWork.EndDate.Date == this.GetManagerWork().EndDate.Date;
            //else 
            return this.DepartmentWork.EndDate.Date == this.GetManagerWork().EndDate.Date;
        }

        /// <summary>
        /// Обработал ли РП запись
        /// </summary>
        public override bool IsChangeStageCanBeComplete
        {
            get
            {
                bool result = false;
                if (this.DepartmentWork.IsEmpty)//если нет связанной работы, то  закрываем
                    result = true;
                else
                    result = IsManagerManuallyDoneDatesCorrection();
                return result;
            }
        }

        public override bool CanBeClosed
        {
            get
            {
                bool result = false;
                if (this.DepartmentWork.IsEmpty)
                    result = true;
                else
                {
                    if (this.StageName == WorkChangeRequest.Stages.DoneByManager && //стадия обработано РП
                     this.ManagerDecision == WorkChangeRequest.ManagerDecisionValue.Accept && // Решение РП = Принять
                     this.CanChangeStageForAcceptedRequestToClose)
                        result = true;
                }
                return result;
            }
        }


        public override ClassObject WorkChangeRequestClass
        {
            get
            {
                return References.Class_WorkChangeRequest_DatesChange;
            }
        }


        /// <summary>
        /// Проверка обработки РП запроса "Изменение сроков"
        /// <para>Рассматриваются два случая, для сроков внутри сроков РП, и вне</para>
        /// </summary>
        /// <param name="workChangeRequest"></param>
        private bool IsManagerManuallyDoneDatesCorrection()
        {
            bool result = false;
            //находим работу в плане подразделения
            ProjectManagementWork departmentWork = this.DepartmentWork;
            //находим работу в плане РП
            ProjectManagementWork managerWork = GetManagerWork();
            if (managerWork == null || managerWork.IsEmpty) throw new ManagerWorkNotFoundException("Отсутствует синхронизированная работа РП для работы с ID=" + departmentWork.ReferenceObject.SystemFields.Id.ToString() + string.Format(". Reference - {0}, ObjectGuid - {1}", this.ReferenceObject.Reference.Name, this.ReferenceObject.SystemFields.Guid.ToString()));//проверяем существование синхронизированной работы РП
            if (OriginalAndAimDatesCoincide)
            {
                //result = managerWork.EndDate.Date > OriginalEndDate.Date;
                result = IsTheWorkEndDateWasIncreasedAfterCreationDateOfCorrectionRequest(managerWork);
            }
            else
            {//даты связанной работы должны быть внутри дат синхронизированной работы РП
                //result = (managerWork.StartDate <= departmentWork.StartDate && managerWork.EndDate >= departmentWork.EndDate);
                //дата окончания связанной работы ранее или равна дате у РП
                result = (managerWork.EndDate >= departmentWork.EndDate);
            }
            return result;
        }

        /// <summary>
        /// Проверяет смещалась ли дата окончания у работы вправо, после создания запроса
        /// </summary>
        /// <param name="managerWork"></param>
        /// <returns></returns>
        private bool IsTheWorkEndDateWasIncreasedAfterCreationDateOfCorrectionRequest(ProjectManagementWork managerWork)
        {
            bool result = false;
            DateTime creationDate = this.CreationDate;
            if (managerWork.ReferenceObject.SystemFields.EditDate >= creationDate)//если работа изменилась после создания запроса
            {
                LogJournal log = WorkLogSearch.Search(managerWork.Guid);

                DateTime? originalDate = log.GetEndDateValue(creationDate);
                if (originalDate != null)//если есть запись об изменении даты окончания, проверяем дальше
                {
                    List<LogRecord> listRecords = log.AllRecords.Where(l => //выбираем записи по условиям
                    l.CreationDate > creationDate && //дата создания записи лога позже создания запроса
                    l.Content.ToLower().Contains(LogJournal.ParameterName.EndDate.ToLower())).ToList(); //запись лога об изменении окончания
                    listRecords = listRecords.Where(l => originalDate < DateTime.Parse(l.NewValue)).ToList();//новая дата из лога позже исходной даты окончания
                    //MessageBox.Show(listRecords.Count.ToString());
                    if (listRecords.Count > 0)//проверяем результат
                    {
                        foreach (var item in listRecords)
                        {
                            //MessageBox.Show(item.CreationDate + "\n" + item.NewValue + ">" + originalDate);
                            if (DateTime.Parse(item.NewValue) > originalDate)//новая дата окончания должна быть позже даты на момент запроса
                            {
                                result = true;
                                break;
                            }
                        }
                    }
                }
                //else throw new ArgumentNullException("Нет записей в логе работы");
            }
            return result;//если работа не изменялась после создания запроса то False
        }

        /// <summary>
        /// Применяем действия по изменению сроков в плане РП
        /// если всё успешно выполнено, возвращает true
        /// </summary>
        internal override bool DoAcceptActions()
        {
            ProjectManagementWork departmentWork = this.DepartmentWork;//находим работу в плане подразделения
            if (departmentWork.IsEmpty) throw new Exception(string.Format("Отсутствует связанная работа в плане подразделения. Reference - {0}, ObjectGuid - {1}", this.ReferenceObject.Reference.Name, this.ReferenceObject.SystemFields.Guid.ToString()));
            //находим работу в плане РП
            ProjectManagementWork managerWork = GetManagerWork();
            if (managerWork == null || managerWork.IsEmpty) throw new Exception(string.Format("Отсутствует синхронизированная работа в плане руководителя проекта. Reference - {0}, ObjectGuid - {1}", this.ReferenceObject.Reference.Name, this.ReferenceObject.SystemFields.Guid.ToString()));
            //if (this.OriginalAndAimDatesCoincide)
            //{
#if !server
            if (ChooseNewStartEndDates(this))
            {
                //изменяем сроки в плане РП на основе выбранных им дат
                managerWork.ChangeDates((DateTime)this.NewStartDate, (DateTime)this.NewEndDate);
            }
            else return false;
#endif
            //}
            //else
            //{
            //    managerWork.ChangeDates(departmentWork.StartDate, departmentWork.EndDate);
            //}
            return true;
        }

        internal override void DoRejectActions()
        {
            //throw new NotImplementedException();
        }

        /// <summary>
        /// Действия необходимые для завершения запроса "Изменение сроков"
        /// <para>При принятии запроса 1. Создание пустой записи в ЖРП</para>
        /// <para>При отклонении запроса 1. Только в ручном режиме</para>
        /// </summary>
        public override bool TryAutomaticallyDoCloseActions()
        {
            //ProjectManagementWork departmentWork = this.DepartmentWork;//находим работу в плане подразделения
            if (this.DepartmentWork.IsEmpty) //если работа существует
            {
                Close(CompleteNote.MissedWork);
                return true;
            }

            if (this.ManagerDecision == ManagerDecisionValue.Accept)//решение РП - принять запрос
            {
                ProjectManagementWork departmentWork = this.DepartmentWork;//находим работу в плане подразделения
                if (departmentWork.IsEmpty) throw new Exception(string.Format("Отсутствует связанная работа в плане подразделения. Reference - {0}, ObjectGuid - {1}", this.ReferenceObject.Reference.Name, this.ReferenceObject.SystemFields.Guid.ToString()));
                //находим работу в плане РП
                ProjectManagementWork managerWork = GetManagerWork();
                if (managerWork == null || managerWork.IsEmpty) throw new Exception(string.Format("Отсутствует синхронизированная работа в плане руководителя проекта. Reference - {0}, ObjectGuid - {1}", this.ReferenceObject.Reference.Name, this.ReferenceObject.SystemFields.Guid.ToString()));

                if (this.OriginalAndAimDatesCoincide)// если даты работ совпадали на момент создания
                {

                    //изменяем сроки в плане подразделения на основе работы РП
                    departmentWork.ChangeDates(managerWork.StartDate, managerWork.EndDate);

                    //return true;
                }
                else if (departmentWork.EndDate.Date < managerWork.EndDate)
                {
                    //изменяем сроки в плане подразделения на основе работы РП
                    departmentWork.ChangeDates(managerWork.StartDate, managerWork.EndDate);
                }
                //в ЖРП связанной работы добавляется пустая запись
                CompleteWithNewEmptyEditReasonNote();//ничего делать не нужно сроки синхронизированы на предыдущем этапе (принятия решения РП)
                //}
                return true;
                //стадия запроса меняется на «Закрыт» автоматически по созданию новой записи с типом "Не определено"
            }
            else if (this.ManagerDecision == ManagerDecisionValue.Reject)
            {
                return false;
            }
            return false;
        }

#if !server
        private bool ChooseNewStartEndDates(WorkChangeRequest_DatesChange changeDates)
        {
            //MacroContext mc = new MacroContext(References.Connection);
            //UIMacroContext uIMacroContext = mc as UIMacroContext;
            //var dialog = mc.CreateInputDialog();
            UIMacroContext uIMacroContext = new UIMacroContext(new MacroContext(References.Connection), null, null, null, null);
            var dial = uIMacroContext.CreateInputDialog();
            //var dialog = ObjectCreator.CreateObject<IInputDialog>();
            var dialog = uIMacroContext.CreateInputDialog();
            ProjectManagementWork departmentWork = changeDates.DepartmentWork;
            ProjectManagementWork managerWork = GetManagerWork();
            if (managerWork == null || managerWork.IsEmpty) throw new Exception(string.Format("Отсутствует синхронизированная работа в плане руководителя проекта. Reference - {0}, ObjectGuid - {1}", this.ReferenceObject.Reference.Name, this.ReferenceObject.SystemFields.Guid.ToString()));

            dialog.Caption = "Выберите новые сроки работы";
            dialog.AddText(changeDates.Content, 1);
            dialog.AddGroup("Работа в плане подразделения");
            dialog.AddDateField("Начало ", departmentWork.StartDate, true, (int)DateTimePickerFormat.Long, "");
            dialog.SetElementEnabled("Начало ", false);
            dialog.AddDateField("Окончание ", departmentWork.EndDate, true, (int)DateTimePickerFormat.Long, "");
            dialog.SetElementEnabled("Окончание ", false);
            dialog.AddText("", 1);
            dialog.AddGroup("Работа в плане РП");
            dialog.AddDateField("Начало  ", managerWork.StartDate, true, (int)DateTimePickerFormat.Long, "");
            dialog.SetElementEnabled("Начало  ", false);
            dialog.AddDateField("Окончание  ", managerWork.EndDate, true, (int)DateTimePickerFormat.Long, "");
            dialog.SetElementEnabled("Окончание  ", false);
            dialog.AddText("", 1);
            //Filter filter = null;
            dialog.AddGroup("Новые даты");

            //устанавливаем предлагаемые значения для РП
            DateTime startDate = departmentWork.StartDate;
            if (managerWork.StartDate > startDate) startDate = managerWork.StartDate;
            DateTime endDate = departmentWork.EndDate;
            if (managerWork.EndDate > endDate) endDate = managerWork.EndDate;

            dialog.AddDateField("Начало", startDate, true, (int)DateTimePickerFormat.Long, "");
            if (departmentWork.Progress > 0)
            {

                dialog.SetElementEnabled("Начало", false);
            }
            dialog.AddDateField("Окончание", endDate, true, (int)DateTimePickerFormat.Long, "");
            dialog.AddText("", 1);

            if (dialog.Show(uIMacroContext))
            {
                DateTime newStartDate = (DateTime)dialog.GetValue("Начало");
                DateTime newEndDate = (DateTime)dialog.GetValue("Окончание");
                bool validDates = false;
                if (changeDates.OriginalAndAimDatesCoincide)//если сроки совпадали
                {
                    if (newEndDate.Date > changeDates.OriginalEndDate.Date)//новая дата окончания должна быть больше исходной
                        validDates = true;
                    else
                    {
                        MessageBox.Show("Новая дата окончания должна быть позже даты окончания в работе подразделения!");
                    }
                }
                else//если сроки не совпадали
                {
                    if (newEndDate.Date >= departmentWork.EndDate.Date)//новая дата окончания должна быть позже или совпадать с работой ПлОП
                        validDates = true;
                    else
                    {
                        MessageBox.Show("Новая дата окончания, должна быть позже или равна, дате окончания в работе подразделения!");
                    }
                }
                if (validDates)
                {
                    changeDates.NewStartDate = newStartDate;
                    changeDates.NewEndDate = newEndDate;
                }
                return validDates;

            }
            return false;
        }
#endif
        /// <summary>
        /// Вызывается при первоначальном создании объекта
        /// </summary>
        protected override void SaveChildOnCreatedParameters()
        {
            this._OriginalStartDate = this.DepartmentWork.StartDate;
            this.ReferenceObject[param_OriginalStartDate_Guid].Value = _OriginalStartDate;
            this._OriginalEndDate = this.DepartmentWork.EndDate;
            this.ReferenceObject[param_OriginalEndDate_Guid].Value = _OriginalEndDate;

            //проверяем входят ли сроки ПлОП в сроки РП
            ProjectManagementWork managerWork = GetManagerWork();
            if (managerWork != null && !managerWork.IsEmpty)
            {
                this._OriginalAndAimDatesCoincide = (this.DepartmentWork.StartDate.Date >= managerWork.StartDate.Date &&
                    this.DepartmentWork.EndDate.Date <= managerWork.EndDate.Date);

                this.ReferenceObject[param_OriginalAndAimDatesCoincide_Guid].Value = OriginalAndAimDatesCoincide;
            }
            else this.ReferenceObject[param_OriginalAndAimDatesCoincide_Guid].Value = false;
        }

        public DateTime? NewStartDate { get; private set; }
        public DateTime? NewEndDate { get; private set; }


        DateTime? _OriginalStartDate;
        DateTime OriginalStartDate
        {
            get
            {
                if (_OriginalStartDate == null)
                {
                    _OriginalStartDate = this.ReferenceObject[param_OriginalStartDate_Guid].GetDateTime();
                    //if (temp != DateTime.MinValue) _OriginalStartDate = temp;
                }
                return ((DateTime)_OriginalStartDate).Date;
            }
        }

        DateTime? _OriginalEndDate;
        DateTime OriginalEndDate
        {
            get
            {
                if (_OriginalEndDate == null)
                {
                    _OriginalEndDate = this.ReferenceObject[param_OriginalEndDate_Guid].GetDateTime();
                    //if (temp != DateTime.MinValue) _OriginalStartDate = temp;
                }
                return ((DateTime)_OriginalEndDate).Date;
            }
        }

        bool? _OriginalAndAimDatesCoincide;
        /// <summary>
        /// Не выходят ли сроки ПлОП работы за сроки работы РП в момент создания запроса
        /// </summary>
        bool OriginalAndAimDatesCoincide
        {
            get
            {
                if (_OriginalAndAimDatesCoincide == null)
                {
                    _OriginalAndAimDatesCoincide = this.ReferenceObject[param_OriginalAndAimDatesCoincide_Guid].GetBoolean();
                    //if (temp != DateTime.MinValue) _OriginalStartDate = temp;
                }
                return (bool)_OriginalAndAimDatesCoincide;
            }
        }

        static readonly Guid param_OriginalStartDate_Guid = new Guid("f41da42f-d54b-4d47-86de-0cc487dc1d3a"); //Guid параметра "Исходная дата начала" Дата    Field_6019
        static readonly Guid param_OriginalEndDate_Guid = new Guid("9870d951-e574-4e62-a315-076b295b6af5"); //Guid параметра "Исходная дата окончания" Дата    Field_6020
        static readonly Guid param_OriginalAndAimDatesCoincide_Guid = new Guid("a4b15000-7c4d-45ec-96c0-86b46ece2b84"); //Guid параметра "Связанные сроки равны или входят в синхронизированные" Параметр "Да/Нет"	Field_6021
    }

    /// <summary>
    /// Запрос на изменение наименования
    /// </summary>
    class WorkChangeRequest_WorkNameChange : WorkChangeRequest
    {
        public WorkChangeRequest_WorkNameChange(EditReasonNote note) : base(note)
        { }
        public WorkChangeRequest_WorkNameChange(ReferenceObject referenceObject) : base(referenceObject)
        { }

        public override bool CanChangeStageForAcceptedRequestToClose { get { return IsDepartmentManuallyChangeWorkName(); } }

        private bool IsDepartmentManuallyChangeWorkName()
        {
            //всегда true т.к. ПлОП первым меняет наименование
            return true;
        }
        public override bool IsChangeStageCanBeComplete
        {
            get
            {
                bool result = false;
                if (this.DepartmentWork.IsEmpty)//если нет связанной работы, то  закрываем
                    result = true;
                else
                    result = IsManagerManuallyChangeWorkName();//проверяем совпадают ли названия у РП и ПлОП
                return result;
            }
        }

        public override ClassObject WorkChangeRequestClass
        {
            get
            {
                return References.Class_WorkChangeRequest_WorkNameChange;
            }
        }

        public override bool CanBeClosed
        {
            get
            {
                bool result = false;
                if (this.DepartmentWork.IsEmpty) result = true;
                else
                {
                    if (this.StageName == WorkChangeRequest.Stages.DoneByManager && //стадия обработано РП
                     this.ManagerDecision == WorkChangeRequest.ManagerDecisionValue.Accept && // Решение РП = Принять
                     this.CanChangeStageForAcceptedRequestToClose) result = true;
                }
                return result;
            }
        }

        /// <summary>
        /// Проверка обработки РП запроса "Изменение наименования"
        /// <para>1. Наименование связанной работы должно совпадать с синхронизированной работой РП</para>
        /// </summary>
        /// <param name="workChangeRequest"></param>
        private bool IsManagerManuallyChangeWorkName()
        {
            ProjectManagementWork departmentWork = this.DepartmentWork;//находим работу в плане подразделения
            //находим работу в плане РП
            ProjectManagementWork managerWork = GetManagerWork();
            if (managerWork == null || managerWork.IsEmpty) throw new ManagerWorkNotFoundException(string.Format("Отсутствует синхронизированная работа в плане руководителя проекта. Reference - {0}, ObjectGuid - {1}", this.ReferenceObject.Reference.Name, this.ReferenceObject.SystemFields.Guid.ToString()));
            return (managerWork.Name == departmentWork.Name);//сравниваем наименования у работ
        }

        /// <summary>
        /// Применяем действия по изменению наименования работы в плане РП
        /// </summary>
        /// <param name="workChangeRequest"></param>
        internal override bool DoAcceptActions()
        {
            ProjectManagementWork departmentWork = this.DepartmentWork;//находим работу в плане подразделения
            if (departmentWork.IsEmpty) throw new ManagerWorkNotFoundException("Отсутствует связанная работа в плане подразделения");
            //находим работу в плане РП
            ProjectManagementWork managerWork = GetManagerWork();
            if (managerWork == null || managerWork.IsEmpty) throw new ManagerWorkNotFoundException(string.Format("Отсутствует синхронизированная работа в плане руководителя проекта. Reference - {0}, ObjectGuid - {1}", this.ReferenceObject.Reference.Name, this.ReferenceObject.SystemFields.Guid.ToString()));
            //изменяем сроки в плане РП на основе работы подразделения
            managerWork.ChangeWorkName(departmentWork.Name);
            return true;
        }
        internal override void DoRejectActions()
        {
            //throw new NotImplementedException();
        }

        /// <summary>
        /// Действия необходимые для завершения запроса
        /// <para>При принятии запроса 1. Создание пустой записи в ЖРП</para>
        /// <para>При отклонении запроса 1. Только в ручном режиме</para>
        /// </summary>
        public override bool TryAutomaticallyDoCloseActions()
        {
            if (this.DepartmentWork.IsEmpty) //если работа существует
            {
                Close(CompleteNote.MissedWork);
                return true;
            }
            if (this.ManagerDecision == ManagerDecisionValue.Accept)
            {
                //в ЖРП связанной работы добавляется пустая запись
                CompleteWithNewEmptyEditReasonNote();
                //стадия запроса меняется на «Закрыт» автоматически по созданию новой записи с типом "Не определено"
            }
            else if (this.ManagerDecision == ManagerDecisionValue.Reject)
            {
                return false;
            }
            return false;
        }

        protected override void SaveChildOnCreatedParameters()
        {
            //нечего записывать
        }
    }

    /// <summary>
    /// Запрос на изменение исполнителя
    /// </summary>
    class WorkChangeRequest_ChangeResource : WorkChangeRequest
    {
        public WorkChangeRequest_ChangeResource(EditReasonNote note) : base(note)
        { }
        public WorkChangeRequest_ChangeResource(ReferenceObject referenceObject) : base(referenceObject)
        { }

        public override bool CanChangeStageForAcceptedRequestToClose { get { return IsDepartmentManuallyDeleteWork(); } }

        private bool IsDepartmentManuallyDeleteWork()
        {
            ProjectManagementWork departmentWork = this.DepartmentWork;//находим работу в плане подразделения
            return departmentWork.IsEmpty; // проверяем её на существование работа должна быть удалена из плана подразделения
        }
        public override bool IsChangeStageCanBeComplete
        {
            get
            {
                bool result = false;
                if (this.DepartmentWork.IsEmpty)//если нет связанной работы, то  закрываем
                    result = true;
                else
                    result = IsManagerManuallyChangeResource();
                return result;
            }
        }

        private bool IsManagerManuallyChangeResource()
        {
            ProjectManagementWork departmentWork = this.DepartmentWork;//находим работу в плане подразделения
            //находим работу в плане РП
            ProjectManagementWork managerWork = GetManagerWork();
            return (managerWork == null || managerWork.IsEmpty);//проверяем существование синхронизированной работы РП её быть не должно
        }

        public override bool CanBeClosed
        {
            get
            {
                bool result = false;
                if (this.DepartmentWork.IsEmpty) result = true;
                else
                {
                    if (this.StageName == WorkChangeRequest.Stages.DoneByManager && //стадия обработано РП
                     this.ManagerDecision == WorkChangeRequest.ManagerDecisionValue.Accept && // Решение РП = Принять
                     this.CanChangeStageForAcceptedRequestToClose) result = true;
                }
                return result;
            }
        }

        internal override bool DoAcceptActions()
        {
            ProjectManagementWork departmentWork = this.DepartmentWork;
            if (!departmentWork.IsEmpty && NewResource != null)
            {
                //ReferenceObject newResource = null;
                //int newLabourHours = 0;
                //#if !server
                //                //выбираем новый ресурс и трудозатраты
                //                ChooseNewResource(ref newResource, ref newLabourHours);
                //#endif
                if (this.NewResource == null)
                {
                    MessageBox.Show("Не выбран новый исполнитель.", "Предупреждение!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                    //throw new NotImplementedException("Не выбран новый исполнитель!");
                }

                //удаляем синхронизацию
                List<ProjectManagementWork> listMasterWorks = Synchronization.GetSynchronizedWorksFromSpace(departmentWork, PlanningSpaceGuidString.WorkingPlans, true);
                if (listMasterWorks.Count > 0)
                {
                    if (listMasterWorks.Count > 1) { throw new InvalidOperationException(string.Format("Укрупнений больше одного!. Reference - {0}, ObjectGuid - {1}", this.ReferenceObject.Reference.Name, this.ReferenceObject.SystemFields.Guid.ToString())); }

                    ProjectManagementWork masterWork = listMasterWorks.First();

                    //в синхронизированной работе РП удаляем ресурсы и добавляем выбранный
                    foreach (var item in masterWork.PlannedNonConsumableResources)
                    {
                        item.Delete();
                    }
                    UsedResource newUsedResource = new UsedResource(this.NewResource, this.NewLabourHours, masterWork.StartDate, masterWork.EndDate);
                    masterWork.AddResource(newUsedResource);

                    Synchronization.DeleteSynchronisationBetween(masterWork, departmentWork);//удаляем синхронизацию
                    return true;
                }
            }
            return false;
            //else throw new ArgumentNullException("Отсутствует работа в плане подразделения, связанная с запросом!");
        }


        public ReferenceObject NewResource { get; set; }
        public int NewLabourHours { get; set; }

        public override ClassObject WorkChangeRequestClass
        {
            get
            {
                return References.Class_WorkChangeRequest_ChangeResource;
            }
        }


        //#if !server
        //        private void ChooseNewResource(ref ReferenceObject newResource, ref int newLabourHours)
        //        {

        //            MacroContext mc = new MacroContext(References.Connection);
        //            UIMacroContext uIMacroContext = new UIMacroContext(mc, null, null, null, null);
        //            MessageBox.Show("1");
        //            mc.Register(uIMacroContext);
        //            MessageBox.Show("2");
        //            //mc.CreateInputDialog();
        //            //var dialog = ObjectCreator.CreateObject<IInputDialog>();
        //            InputDialog dialog = new InputDialog(uIMacroContext, "Выберите нового исполнителя");
        //            dialog.Caption = "Выберите нового исполнителя";
        //            Filter filter = null;
        //            dialog.AddSelectFromReference("Выберите исполнителя", "Ресурсы", "", newResource, true, "");
        //            dialog.AddInteger("Трудозатраты", newLabourHours, true);

        //            if (dialog.Show())
        //            {
        //#if test
        //        MessageBox.Show(dialog.GetValue("Выберите исполнителя").ToString());
        //#endif

        //                newResource = References.Resources.Find((int)(dialog.GetValue("Выберите исполнителя")));
        //                newLabourHours = dialog.GetValue("Трудозатраты");
        //            }
        //        }
        //#endif
        internal override void DoRejectActions()
        {
            //throw new NotImplementedException();
        }

        public override bool TryAutomaticallyDoCloseActions()
        {
            if (this.ManagerDecision == ManagerDecisionValue.Accept)
            {
                ProjectManagementWork departmentWork = this.DepartmentWork;//находим работу в плане подразделения
                if (!departmentWork.IsEmpty) //если работа существует
                {
                    departmentWork.Delete();//удаляем работу
                }
                Close(CompleteNote.AutoClose);
                return true;
            }
            else if (this.ManagerDecision == ManagerDecisionValue.Reject)
            {
                return false;
            }
            return false;
        }

        protected override void SaveChildOnCreatedParameters()
        {
            //нечего записывать
        }
    }

    /// <summary>
    /// Запрос на удаление работы
    /// </summary>
    class WorkChangeRequest_DeleteWork : WorkChangeRequest
    {
        public WorkChangeRequest_DeleteWork(EditReasonNote note) : base(note)
        { }
        public WorkChangeRequest_DeleteWork(ReferenceObject referenceObject) : base(referenceObject)
        { }

        public override bool CanChangeStageForAcceptedRequestToClose { get { return IsDepartmentManuallyDeleteWork(); } }

        private bool IsDepartmentManuallyDeleteWork()
        {
            ProjectManagementWork departmentWork = this.DepartmentWork;//находим работу в плане подразделения
            return departmentWork.IsEmpty; // проверяем её на существование
        }

        /// <summary>
        /// Выполнены ли все условия для закрытия запроса
        /// </summary>
        public override bool IsChangeStageCanBeComplete
        {
            get
            {
                bool result = this.DepartmentWork.IsEmpty || IsManagerManuallyDeleteWork();
                return result;
            }
        }

        public override bool CanBeClosed
        {
            get
            {
                bool result = false;
                if (this.DepartmentWork.IsEmpty) result = true;
                else
                {
                    if (this.StageName == WorkChangeRequest.Stages.DoneByManager && //стадия обработано РП
                     this.ManagerDecision == WorkChangeRequest.ManagerDecisionValue.Accept && // Решение РП = Принять
                     this.CanChangeStageForAcceptedRequestToClose) result = true;
                }
                return result;
            }
        }

        public override ClassObject WorkChangeRequestClass
        {
            get
            {
                return References.Class_WorkChangeRequest_DeleteWork;
            }
        }

        private bool IsManagerManuallyDeleteWork()
        {
            ProjectManagementWork departmentWork = this.DepartmentWork;//находим работу в плане подразделения
            //if (departmentWork.IsEmpty) throw new Exception("Отсутствует связанная работа в плане подразделения");
            //находим работу в плане РП
            ProjectManagementWork managerWork = GetManagerWork();
            return (managerWork == null || managerWork.IsEmpty);//проверяем отсутствие работы
        }

        /// <summary>
        /// <para>Действия необходимые для обработки РП запроса "Удаление работы"</para>
        /// 1. Удаление синхронизированной работы РП
        /// </summary>
        /// <param name="workChangeRequest"></param>
        internal override bool DoAcceptActions()
        {
            ProjectManagementWork departmentWork = this.DepartmentWork;//находим работу в плане подразделения
            if (departmentWork.IsEmpty) throw new Exception("Отсутствует связанная работа в плане подразделения" + string.Format(". Reference - {0}, ObjectGuid - {1}", this.ReferenceObject.Reference.Name, this.ReferenceObject.SystemFields.Guid.ToString()));//проверяем существование связанной работы
                                                                                                                                                                                                                                                                 //находим работу в плане РП
            ProjectManagementWork managerWork = GetManagerWork();
            if (managerWork != null && !managerWork.IsEmpty) //если работа существует
            {
                managerWork.Delete();//удаляем работу
            }
            return true;
        }

        internal override void DoRejectActions()
        {
            //throw new NotImplementedException();
        }

        /// <summary>
        /// <para> Выполняет действия необходимые для завершения запроса "Удаление работы"</para>
        /// 1. Удаление связанной работы
        /// </summary>
        /// <returns>true если все действия выполнены; false - если требуются действия пользователя</returns>
        public override bool TryAutomaticallyDoCloseActions()
        {
            if (this.ManagerDecision == ManagerDecisionValue.Accept)
            {
                ProjectManagementWork departmentWork = this.DepartmentWork;//находим работу в плане подразделения
                if (!departmentWork.IsEmpty) //если работа существует
                {
                    departmentWork.Delete();//удаляем работу
                }
                Close(CompleteNote.AutoClose);
                return true;
            }
            else if (this.ManagerDecision == ManagerDecisionValue.Reject)
            {
                return false;
            }
            return false;
        }

        protected override void SaveChildOnCreatedParameters()
        {
            //нечего записывать
        }
    }

    /// <summary>
    /// Запрос на добавление работы
    /// </summary>
    class WorkChangeRequest_AddWork : WorkChangeRequest
    {
        public WorkChangeRequest_AddWork(EditReasonNote note) : base(note)
        { }
        public WorkChangeRequest_AddWork(ReferenceObject referenceObject) : base(referenceObject)
        { }

        public override bool CanChangeStageForAcceptedRequestToClose { get { return IsDepartmentManuallyAddWork(); } }

        private bool IsDepartmentManuallyAddWork()
        {
            //всегда true т.к. ПлОП первым добавляет работу
            return true;
        }
        public override bool IsChangeStageCanBeComplete
        {
            get
            {
                bool result = false;
                if (this.DepartmentWork.IsEmpty)//если нет связанной работы, то  закрываем
                    result = true;
                else
                    result = IsManagerManuallyAddWork();
                return result;
            }
        }

        public override ClassObject WorkChangeRequestClass
        {
            get
            {
                return References.Class_WorkChangeRequest_AddWork;
            }
        }

        public override bool CanBeClosed
        {
            get
            {
                bool result = false;
                if (this.DepartmentWork.IsEmpty) result = true;
                else
                {
                    if (this.StageName == WorkChangeRequest.Stages.DoneByManager && //стадия обработано РП
                     this.ManagerDecision == WorkChangeRequest.ManagerDecisionValue.Accept && // Решение РП = Принять
                     this.CanChangeStageForAcceptedRequestToClose) result = true;
                }
                return result;
            }
        }

        private bool IsManagerManuallyAddWork()
        {
            ProjectManagementWork departmentWork = this.DepartmentWork;//находим работу в плане подразделения
            //находим работу в плане РП
            ProjectManagementWork managerWork = GetManagerWork();
            return (managerWork != null && !managerWork.IsEmpty);//проверяем существование работы
        }

        internal override bool DoAcceptActions()
        {
            ProjectManagementWork departmentWork = this.DepartmentWork;//находим работу в плане подразделения
            if (departmentWork.IsEmpty) throw new Exception("Отсутствует связанная работа в плане подразделения");//проверяем существование связанной работы
                                                                                                                  //находим работу в плане РП
            ProjectManagementWork managerWork = GetManagerWork();
            if (managerWork != null && !managerWork.IsEmpty) //если работа существует
            {
                throw new Exception("Синхронизированная работа РП уже существует" + string.Format(". Reference - {0}, ObjectGuid - {1}", this.ReferenceObject.Reference.Name, this.ReferenceObject.SystemFields.Guid.ToString()));
            }
            else
            {
                //копируем связанную работу в план РП
                managerWork = CreateCopyInManagerProject(departmentWork);
                //Синхронизируем работы
                Synchronization.SyncronizeWorks(managerWork, departmentWork);
                КопироватьИспользуемыеРесурсы(departmentWork, managerWork);
                return true;
            }
        }

        private ProjectManagementWork CreateCopyInManagerProject(ProjectManagementWork departmentWork)
        {
            ProjectManagementWork parentDepartmentWork = departmentWork.ParentWork;
            if (parentDepartmentWork == null || parentDepartmentWork.IsEmpty) throw new Exception("Нет синхронизированной работы у родительской работы - " + departmentWork.Name + string.Format(". Reference - {0}, ObjectGuid - {1}", this.ReferenceObject.Reference.Name, this.ReferenceObject.SystemFields.Guid.ToString()));
            //находим родительскую работу в плане РП
            ProjectManagementWork parentManagerWork = GetManagerWork();
            if (parentManagerWork == null || parentManagerWork.IsEmpty) throw new Exception("Нет синхронизированной работы у родительской работы - " + departmentWork.Name + string.Format(". Reference - {0}, ObjectGuid - {1}", this.ReferenceObject.Reference.Name, this.ReferenceObject.SystemFields.Guid.ToString()));
            ProjectManagementWork newManagerWork = CreateChildCopy(parentManagerWork, departmentWork);
            return newManagerWork;
        }

        /// <summary>
        /// Копирует все ресурсы из одной работы в другую
        /// </summary>
        /// <param name="copyFrom"></param>
        /// <param name="copyTo"></param>
        private void КопироватьИспользуемыеРесурсы(ProjectManagementWork copyFrom, ProjectManagementWork copyTo)
        {
            foreach (var resource in copyFrom.ReferenceObject.GetObjects(ProjectManagementWork.PM_link_UsedResources_GUID))
            {
                ReferenceObject copyResource = resource.CreateCopy();
                //copyResource.Refresh(resource);
                copyResource.SetLinkedObject(ProjectManagementWork.PM_link_UsedResources_GUID, copyTo.ReferenceObject);
                copyResource.EndChanges();
                copyResource = null;
            }
        }

        /// <summary>
        /// Создаёт копию работы в указанном родителе
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="originForCopy"></param>
        /// <returns></returns>
        private ProjectManagementWork CreateChildCopy(ProjectManagementWork parent, ProjectManagementWork originForCopy)
        {

            ReferenceObject исходныйОбъект = originForCopy.ReferenceObject;
            ReferenceObject целевойОбъект = parent.ReferenceObject;

            ReferenceObject копияИсходногоОбъекта = References.ProjectManagementReference.CreateReferenceObject(целевойОбъект, References.Class_ProjectManagementWork);
            foreach (Parameter parameter in копияИсходногоОбъекта.ParameterValues)
            {
                if (parameter.IsReadOnly) continue;
                if (ProjectManagementWork.listParametersToSkip.Contains(parameter.ParameterInfo.Guid) || ProjectManagementWork.listLinksToSkip.Contains(parameter.ParameterInfo.Guid)) continue;

                копияИсходногоОбъекта[parameter.ParameterInfo].Value = исходныйОбъект[parameter.ParameterInfo].Value;
            }

            копияИсходногоОбъекта.SetParent(целевойОбъект);

            копияИсходногоОбъекта.EndChanges();

            if (копияИсходногоОбъекта == null)
                throw new Exception("Не удалось создать копию в целевом объекте - " + целевойОбъект.ToString() + string.Format(". Reference - {0}, ObjectGuid - {1}", this.ReferenceObject.Reference.Name, this.ReferenceObject.SystemFields.Guid.ToString()));

            return new ProjectManagementWork(копияИсходногоОбъекта);

        }

        internal override void DoRejectActions()
        {
            //throw new NotImplementedException();
        }

        /// <summary>
        /// Действия необходимые для завершения запроса "Изменение сроков"
        /// <para>1. Создание пустой записи в ЖРП</para>
        /// </summary>
        public override bool TryAutomaticallyDoCloseActions()
        {
            bool isClosed = false;
            if (this.DepartmentWork.IsEmpty) //если работа существует
            {
                Close(CompleteNote.MissedWork);
                isClosed = true;
            }
            else
            {
                if (this.ManagerDecision == ManagerDecisionValue.Accept)
                {
                    //в ЖРП связанной работы добавляется пустая запись
                    CompleteWithNewEmptyEditReasonNote();
                    //стадия запроса меняется на «Закрыт» автоматически по созданию новой записи с типом "Не определено"
                    isClosed = true;
                }
                else if (this.ManagerDecision == ManagerDecisionValue.Reject)
                {
                    isClosed = false;
                }
            }
            return isClosed;
        }

        protected override void SaveChildOnCreatedParameters()
        {
            //throw new NotImplementedException();
        }
    }

    internal class References
    {
        // Справочник "Журнал редактирования планов"
        static readonly Guid ref_EditReasonJournal_Guid = new Guid("be37cb0f-4c5c-4783-908a-fe3105fff643"); // Справочник "Журнал редактирования планов"
                                                                                                            //Справочник "Управление проектами"
        static readonly Guid ref_ProjectManagement_Guid = new Guid("86ef25d4-f431-4607-be15-f52b551022ff"); // Справочник "Управление проектами"



        // Справочник "Запросы коррекции работ"
        static readonly Guid WorkChangeRequests_ref_Guid = new Guid("9387cd96-d3cf-4cb6-997d-c4f4e59f8a21"); // Справочник "Запросы коррекции работ"
        static readonly Guid CR_class_PMWorkChangeRequest_Guid = new Guid("f79a353c-00da-4248-bd4f-dab6543f95b0"); // Тип "Запросы коррекции работ"

        //поля
        private static Reference _WorkChangeRequestReference;
        private static Reference _ProjectManagementReference;
        private static ReferenceInfo _ProjectManagementReferenceInfo;
        private static ReferenceInfo _WorkChangeRequestReferenceInfo;

        public static ServerConnection Connection
        {
            get { return ServerGateway.Connection; }
        }

        private static Reference GetReference(ref Reference reference, ReferenceInfo referenceInfo)
        {
            if (reference == null)
                reference = referenceInfo.CreateReference();


            return reference;
        }


        static UserReference _UsersReference;
        /// <summary>
        /// Справочник "Группы и пользователи"
        /// </summary>
        public static UserReference UsersReference
        {
            get
            {
                if (_UsersReference == null)
                    _UsersReference = new UserReference(Connection);

                return _UsersReference;
            }
        }

        /// <summary>
        /// Справочник "Запросы коррекции работ"
        /// </summary>
        public static Reference WorkChangeRequestReference
        {
            get
            {
                if (_WorkChangeRequestReference == null)
                    return GetReference(ref _WorkChangeRequestReference, WorkChangeRequestReferenceInfo);
                _WorkChangeRequestReference.Refresh();
                return _WorkChangeRequestReference;
            }
        }

        private static Reference _WorkLogReference;
        /// <summary>
        /// Справочник "Log"
        /// </summary>
        public static Reference WorkLogReference
        {
            get
            {
                if (_WorkLogReference == null)
                    return GetReference(ref _WorkLogReference, WorkLogReferenceInfo);

                return _WorkLogReference;
            }
        }

        private static ClassObject _Class_PMWorkChangeRequest;

        /// <summary>
        /// класс Запрос коррекции работ
        /// </summary>
        public static ClassObject Class_WorkChangeRequest
        {
            get
            {
                if (_Class_PMWorkChangeRequest == null)
                    _Class_PMWorkChangeRequest = WorkChangeRequestReference.Classes.Find(CR_class_PMWorkChangeRequest_Guid);

                return _Class_PMWorkChangeRequest;
            }
        }

        private static ClassObject _Class_WorkChangeRequest_DatesChange;

        /// <summary>
        /// класс Запрос коррекции работ изменение сроков
        /// </summary>
        public static ClassObject Class_WorkChangeRequest_DatesChange
        {
            get
            {
                if (_Class_WorkChangeRequest_DatesChange == null)
                {
                    Guid class_WorkChangeRequest_DatesChange_Guid = new Guid("c297e14e-20ef-41a9-afc1-ee08cd4d87b2"); //Guid типа "Перенос сроков"
                    _Class_WorkChangeRequest_DatesChange = WorkChangeRequestReference.Classes.Find(class_WorkChangeRequest_DatesChange_Guid);
                }

                return _Class_WorkChangeRequest_DatesChange;
            }
        }

        private static ClassObject _Class_WorkChangeRequest_WorkNameChange;

        /// <summary>
        /// класс Запрос коррекции работ изменение имени
        /// </summary>
        public static ClassObject Class_WorkChangeRequest_WorkNameChange
        {
            get
            {
                if (_Class_WorkChangeRequest_WorkNameChange == null)
                {
                    Guid class_WorkChangeRequest_WorkNameChange_Guid = new Guid("445652d5-1c9c-4ec9-a00c-9e28e91a3cb5"); //Guid типа "Изменение наименования"
                    _Class_WorkChangeRequest_WorkNameChange = WorkChangeRequestReference.Classes.Find(class_WorkChangeRequest_WorkNameChange_Guid);
                }

                return _Class_WorkChangeRequest_WorkNameChange;
            }
        }

        private static ClassObject _Class_WorkChangeRequest_ChangeResource;

        /// <summary>
        /// класс Запрос коррекции работ изменение исполнителя
        /// </summary>
        public static ClassObject Class_WorkChangeRequest_ChangeResource
        {
            get
            {
                if (_Class_WorkChangeRequest_ChangeResource == null)
                {
                    Guid class_ChangeResource_Guid = new Guid("28a6872d-1fb0-4bac-88e9-546ea1e4e3dd"); //Guid типа "Изменение исполнителя"
                    _Class_WorkChangeRequest_ChangeResource = WorkChangeRequestReference.Classes.Find(class_ChangeResource_Guid);
                }

                return _Class_WorkChangeRequest_ChangeResource;
            }
        }

        private static ClassObject _Class_WorkChangeRequest_DeleteWork;

        /// <summary>
        /// класс Запрос коррекции работ Удаление работы
        /// </summary>
        public static ClassObject Class_WorkChangeRequest_DeleteWork
        {
            get
            {
                if (_Class_WorkChangeRequest_DeleteWork == null)
                {
                    Guid class_DeleteWork_Guid = new Guid("e959a9e7-e8ef-402b-950f-6d3f14521908"); //Guid типа "Удаление работы"
                    _Class_WorkChangeRequest_DeleteWork = WorkChangeRequestReference.Classes.Find(class_DeleteWork_Guid);
                }

                return _Class_WorkChangeRequest_DeleteWork;
            }
        }

        private static ClassObject _Class_WorkChangeRequest_AddWork;

        /// <summary>
        /// класс Запрос коррекции работ Добавление работы
        /// </summary>
        public static ClassObject Class_WorkChangeRequest_AddWork
        {
            get
            {
                if (_Class_WorkChangeRequest_AddWork == null)
                {
                    Guid class_AddWork_Guid = new Guid("8a745d90-aeb4-4089-ad8e-a57d12b86bac"); //Guid типа "Добавление работы"
                    _Class_WorkChangeRequest_AddWork = WorkChangeRequestReference.Classes.Find(class_AddWork_Guid);
                }

                return _Class_WorkChangeRequest_AddWork;
            }
        }

        private static ClassObject _Class_ProjectManagementWork;
        /// <summary>
        /// тип Работа - справочника Управление проектами
        /// </summary>
        public static ClassObject Class_ProjectManagementWork
        {
            get
            {
                if (_Class_ProjectManagementWork == null)
                    _Class_ProjectManagementWork = ProjectManagementReference.Classes.Find(ProjectManagementWork.PM_class_Work_Guid);

                return _Class_ProjectManagementWork;
            }
        }

        /// <summary>
        /// Справочник "Управление проектами"
        /// </summary>
        public static Reference ProjectManagementReference
        {
            get
            {
                if (_ProjectManagementReference == null)
                    _ProjectManagementReference = GetReference(ref _ProjectManagementReference, ProjectManagementReferenceInfo);
                _ProjectManagementReference.Refresh();
                return _ProjectManagementReference;
            }
        }



        /// <summary>
        /// Справочник "Управление проектами"
        /// </summary>
        public static ReferenceInfo ProjectManagementReferenceInfo
        {
            get { return GetReferenceInfo(ref _ProjectManagementReferenceInfo, ref_ProjectManagement_Guid); }
        }

        /// <summary>
        /// Справочник "Запросы коррекции работ"
        /// </summary>
        internal static ReferenceInfo WorkChangeRequestReferenceInfo
        {
            get { return GetReferenceInfo(ref _WorkChangeRequestReferenceInfo, WorkChangeRequests_ref_Guid); }
        }

        static ReferenceInfo _WorkLogReferenceInfo;
        /// <summary>
        /// Справочник "Log"
        /// </summary>
        internal static ReferenceInfo WorkLogReferenceInfo
        {
            get
            {
                Guid WorkLogReference_ref_Guid = new Guid("60b68da9-15e9-4f9d-a46a-d9b02367b154");
                return GetReferenceInfo(ref _WorkLogReferenceInfo, WorkLogReference_ref_Guid);
            }
        }
        #region "Используемые ресурсы"

        // Справочник "Используемые ресурсы"
        static readonly Guid ref_UsedResources_Guid = new Guid("3459a8fb-6bca-47ca-971a-1572b684e92e");        //Guid справочника - "Используемые ресурсы"
        static readonly Guid UR_class_NonConsumableResources_Guid = new Guid("8473c817-68fd-479c-a4c3-d6b3b405ea5d"); //Guid типа "Нерасходуемые ресурсы"

        private static Reference _UsedResourcesReference;
        /// <summary>
        /// Справочник "Используемые ресурсы"
        /// </summary>
        public static Reference UsedResources
        {
            get
            {
                if (_UsedResourcesReference == null)
                    _UsedResourcesReference = GetReference(ref _UsedResourcesReference, UsedResourcesReferenceInfo);

                return _UsedResourcesReference;
            }
        }

        private static ReferenceInfo _UsedResourcesReferenceInfo;
        /// <summary>
        /// Справочник "Используемые ресурсы"
        /// </summary>
        private static ReferenceInfo UsedResourcesReferenceInfo
        {
            get { return GetReferenceInfo(ref _UsedResourcesReferenceInfo, ref_UsedResources_Guid); }
        }

        private static ClassObject _Class_NonConsumableResources;
        /// <summary>
        /// класс Нерасходуемые ресурсы
        /// </summary>
        public static ClassObject Class_NonConsumableResources
        {
            get
            {
                if (_Class_NonConsumableResources == null)
                    _Class_NonConsumableResources = UsedResources.Classes.Find(UR_class_NonConsumableResources_Guid);

                return _Class_NonConsumableResources;
            }
        }
        #endregion

        #region "Ресурсы"

        // Справочник "Используемые ресурсы"
        static readonly Guid ref_Resources_Guid = new Guid("fe80ab68-01e1-4a95-96cf-602ec877ff19");        //Guid справочника - "Ресурсы"
        private static Reference _ResourcesReference;
        /// <summary>
        /// Справочник "Ресурсы"
        /// </summary>
        public static Reference Resources
        {
            get
            {
                if (_ResourcesReference == null)
                    _ResourcesReference = GetReference(ref _ResourcesReference, ResourcesReferenceInfo);

                return _UsedResourcesReference;
            }
        }

        private static ReferenceInfo _ResourcesReferenceInfo;
        /// <summary>
        /// Справочник "Ресурсы"
        /// </summary>
        private static ReferenceInfo ResourcesReferenceInfo
        {
            get { return GetReferenceInfo(ref _ResourcesReferenceInfo, ref_Resources_Guid); }
        }

        //private static ClassObject _Class_NonConsumableResources;
        ///// <summary>
        ///// класс Нерасходуемые ресурсы
        ///// </summary>
        //public static ClassObject Class_NonConsumableResources
        //{
        //    get
        //    {
        //        if (_Class_NonConsumableResources == null)
        //            _Class_NonConsumableResources = UsedResources.Classes.Find(UR_class_NonConsumableResources_Guid);

        //        return _Class_NonConsumableResources;
        //    }
        //}
        #endregion

        private static ReferenceInfo GetReferenceInfo(ref ReferenceInfo referenceInfo, Guid referenceGuid)
        {
            if (referenceInfo == null)
                referenceInfo = Connection.ReferenceCatalog.Find(referenceGuid);

            return referenceInfo;
        }
    }

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

        EditReasonNote.Type? _TypeCode;
        /// <summary>
        /// Код Типа редактирования
        /// </summary>
        public EditReasonNote.Type TypeCode
        {
            get
            {
                if (_TypeCode == null)
                    _TypeCode = (EditReasonNote.Type)ReferenceObject[EditReasonNote.ERN_param_EditType_Guid].GetInt16();
                return (EditReasonNote.Type)_TypeCode;
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
        { get { return TypeNames[(int)TypeCode]; } }

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
        public enum Type
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
        public static Reference Reference
        {
            get
            {
                if (_Reference == null)
                    _Reference = EditReasonNote.ReferenceInfo.CreateReference();
                else _Reference.Refresh();
                return _Reference;
            }
        }

        // поиск типа объекта
        ClassObject classObject
        {
            get { return EditReasonNote.ReferenceInfo.Classes.Find(EditReasonNote.ERN_type_ChangeReason_Guid); }
        }

        public DateTime CreateDate
        {
            get
            {
                if (this.ReferenceObject != null)
                {
                    return this.ReferenceObject.SystemFields.CreationDate;
                }
                return DateTime.MinValue;
            }
        }

        // Справочник "Журнал редактирования планов"
        public static readonly Guid ref_EditReasonNotes_Guid = new Guid("be37cb0f-4c5c-4783-908a-fe3105fff643"); // Справочник "Журнал редактирования планов"

        public static readonly Guid ERN_type_ChangeReason_Guid = new Guid("b6591f71-9e79-4c86-92c5-7a978d8309f1"); // Тип "Причина редактирования"
        static readonly Guid ERN_param_EditType_Guid = new Guid("37e0d2d3-bbcd-43db-9f52-cd30a66b526d"); // Параметр "Тип редактирования"
        static readonly Guid ERN_param_Reason_Guid = new Guid("93e405ca-b265-4671-8229-9f5d6a67e786"); // Параметр "Причина редактирования"
        static readonly Guid ERN_param_NeedActionManager_Guid = new Guid("227f4813-7447-493a-aa4d-d2a6011823fe"); // Параметр "Требуется реакция РП" Параметр "Да/Нет"	Field_6006
        public static readonly Guid ERN_link_ToPrjctMgmntN1_Guid = new Guid("79b01004-3c10-465a-a6fb-fe2aa95ae5b8"); // Связь n:1 "Журнал редактирования планов - Управление проектами"
    }

    /// <summary>
    /// Информационный класс Элемент проекта - Управления проектами
    /// ver 1.11
    /// </summary>
    public class ProjectManagementWork : ProjectTreeItem
    {
        public ProjectManagementWork(string stringGuid)
        {
            Guid = new Guid(stringGuid);
            ReferenceObject work = FindWork(Guid);
            ReferenceObject = work;
            //base.ReferenceObject = work;
        }

        public ProjectManagementWork(ReferenceObject work)
        {
            ReferenceObject = work;
            this.Guid = work.SystemFields.Guid;
        }

        public ProjectManagementWork(Guid guid)
        {
            Guid = guid;
            ReferenceObject work = FindWork(Guid);
            ReferenceObject = work;

        }
        ReferenceObject _ReferenceObject;
        //public ReferenceObject ReferenceObject
        //{
        //    get
        //    {
        //        return _ReferenceObject;
        //    }
        //    set
        //    {
        //        if (_ReferenceObject != value)
        //        {
        //            _ReferenceObject = value;
        //            if (_ReferenceObject != null)
        //            {
        //                LastEditDate = _ReferenceObject.SystemFields.EditDate;
        //            }

        //        }
        //    }
        //}

        public Guid Guid { get; private set; }

        string _Number;
        /// <summary>
        /// номер работы
        /// </summary>
        public string Number
        {
            get
            {
                if (this.IsProject || this.ParentWork == null)
                { _Number = ""; }
                else if (this.ParentWork != null)
                {
                    if (string.IsNullOrWhiteSpace(this.ParentWork.Number))
                        _Number = this.ReferenceObject.SystemFields.Order.ToString();
                    else
                        _Number = this.ParentWork.Number + "." + this.ReferenceObject.SystemFields.Order.ToString();
                }
                return _Number;
            }
        }


        /// <summary>
        /// Находит синхронизированную работу из пространства Рабочие планы двигаясь вверх по дереву от указанной работы
        /// </summary>
        /// <param name="work"></param>
        /// <returns>работа из пространства Рабочие планы, если не найдена то null</returns>
        private ProjectManagementWork GetTheWorkFromWorkingPlanningSpace(ProjectManagementWork work)
        {
            ProjectManagementWork result = null;
            if (work == null) return result;
            //MacroContext mc = new MacroContext(ServerGateway.Connection);
            // MessageBox.Show(work.ToString());
            if (work.PlanningSpace.ToString() == PlanningSpaceGuidString.WorkingPlans)
            {
                //MessageBox.Show("work from PlanningSpaceGuidString" + work.ToString());
                return work;
            }
            else
            {
                foreach (var item in Synchronization.GetSynchronizedWorksFromSpace(work, PlanningSpaceGuidString.WorkingPlans))
                {
                    while (result == null)
                    {
                        result = GetTheWorkFromWorkingPlanningSpace(item);
                    }
                }
            }
            if (result == null && work.ParentWork != null) result = GetTheWorkFromWorkingPlanningSpace(work.ParentWork);
            return result;
        }

        /// <summary>
        /// Находит кубик из 1С, если не получилось то null
        /// </summary>
        /// <param name="planningSpaceWork"></param>
        /// <returns>кубик из 1С, если не получилось то null</returns>
        private ProjectManagementWork GetTheProjectFrom1C(ProjectManagementWork planningSpaceWork)
        {

            //MessageBox.Show("GetTheProjectFrom1C" + planningSpaceWork.PlanningSpace.ToString() + "\nIsProjectFrom1C = " + planningSpaceWork.IsProjectFrom1C.ToString()+"\n" + planningSpaceWork.Name);
            if (planningSpaceWork.PlanningSpace.ToString() != PlanningSpaceGuidString.WorkingPlans)
            {

                //MessageBox.Show(planningSpaceWork.PlanningSpace.ToString() + " != " + PlanningSpaceGuidString.WorkingPlans);
                planningSpaceWork = GetTheWorkFromWorkingPlanningSpace(planningSpaceWork);
            }
            else
            {
                if (planningSpaceWork.IsProjectFrom1C) return planningSpaceWork;
            }
            //if (planningSpaceWork.Class.Guid == PM_class_Project_Guid) return planningSpaceWork;
            if (planningSpaceWork == null)
            {
                //MessageBox.Show("planningSpaceWork == null");
                return null;
            }
            //MessageBox.Show("GetTheParentProjectFrom1C" + planningSpaceWork.PlanningSpace.ToString() + planningSpaceWork.Name);
            return GetTheParentProjectFrom1C(planningSpaceWork);
        }

        private ProjectManagementWork GetTheParentProjectFrom1C(ProjectManagementWork work)
        {
            if (work.IsProjectFrom1C) return work;
            //if (planningSpaceWork.Class.Guid == PM_class_Project_Guid) return planningSpaceWork;
            if (work.ParentWork == null) return null;
            return GetTheParentProjectFrom1C(work.ParentWork);
        }

        private ReferenceObject FindWork(Guid guid)
        {
            Guid PMReferenceGuid = new Guid("86ef25d4-f431-4607-be15-f52b551022ff");
            //Reference PMReference = ServerGateway.Connection.ReferenceCatalog.Find(PMReferenceGuid).CreateReference();
            Reference PMReference = References.ProjectManagementReference;
            ReferenceObject work = PMReference.Find(guid);
            return work;
        }

        string _Status;
        public string Status
        {
            get
            {
                if (_Status == null)
                {
                    MacroContext mc = new MacroContext(Connection);
                    _Status = (string)(mc.RunMacro("20085642-be1b-41cd-8105-e932eea6fa0c", "WorkState", ReferenceObject));
                }
                return _Status;
            }
        }

        string _Name;
        public string Name
        {
            get
            {
                if (_Name == null)
                {
                    if (ReferenceObject != null)
                        _Name = ReferenceObject[PM_param_Name_GUID].GetString();
                    else _Name = "";
                }
                return _Name;
            }
        }

        DateTime? _StartDate;
        public DateTime StartDate
        {
            get
            {
                if (_StartDate == null)
                {
                    if (ReferenceObject != null)
                        _StartDate = ReferenceObject[PM_param_PlanStartDate_GUID].GetDateTime();
                    else _StartDate = DateTime.MinValue;
                }
                return (DateTime)_StartDate;
            }
            private set
            {
                if (_StartDate != value)
                    _StartDate = value;
            }
        }
        DateTime? _EndDate;
        public DateTime EndDate
        {
            get
            {
                if (_EndDate == null)
                {
                    if (ReferenceObject != null)
                        _EndDate = ReferenceObject[PM_param_PlanEndDate_GUID].GetDateTime();
                    else _EndDate = DateTime.MinValue;
                }
                return (DateTime)_EndDate;
            }
            private set
            {
                if (_EndDate != value)
                    _EndDate = value;
            }
        }

        //string _UsedResourcesNames;
        public string UsedResourcesNames
        {
            get
            {
                string result = "";
                if (PlannedNonConsumableResources != null)
                {
                    int i = 0;
                    foreach (var item in PlannedNonConsumableResources)
                    {
                        if (i == 0) result += item.Resource.ToString();
                        else result += Environment.NewLine + item.Resource.ToString();
                        i++;
                    }
                }
                return result;
            }
        }

        List<UsedResource> _PlannedNonConsumableResources;
        public IEnumerable<UsedResource> PlannedNonConsumableResources
        {
            get
            {
                if (_PlannedNonConsumableResources == null)
                {
                    List<ReferenceObject> resources = this.ReferenceObject.GetObjects(ProjectManagementWork.PM_link_UsedResources_GUID).Where(res => (res.Class == References.Class_NonConsumableResources)).ToList();
                    _PlannedNonConsumableResources = new List<UsedResource>();
                    foreach (var item in resources)
                    {
                        UsedResource usedResource = new UsedResource(item);
                        if (usedResource.IsPlanned) _PlannedNonConsumableResources.Add(usedResource);
                    }
                }
                return _PlannedNonConsumableResources;
            }
        }

        double? _LabourCost;
        public double LabourCost
        {
            get
            {
                if (_LabourCost == null)
                {
                    if (ReferenceObject != null)
                    { _LabourCost = this.ReferenceObject[PM_param_SumPlanResourceCount_GUID].GetDouble(); }
                    else _LabourCost = 0;
                }
                return (double)_LabourCost;
            }
        }

        string _ProjectCipherFrom1C;
        public string ProjectCipherFrom1C
        {
            get
            {
                if (_ProjectCipherFrom1C == null && ProjectFrom1C != null)
                {
                    //MessageBox.Show("ProjectFrom1C -" + ProjectFrom1C.ToString());
                    _ProjectCipherFrom1C = ProjectFrom1C.ReferenceObject[PM_param_Cipher_GUID].GetString();
                }
                if (ProjectFrom1C == null)
                {
                    //MessageBox.Show("ProjectFrom1C == null");
                    _ProjectCipherFrom1C = "Не указан";
                }
                return _ProjectCipherFrom1C;
            }
        }

        string _ProjectNumberFrom1C;
        public string ProjectNumberFrom1C
        {
            get
            {
                if (_ProjectNumberFrom1C == null && ProjectFrom1C != null)
                {
                    //MessageBox.Show("ProjectFrom1C -" + ProjectFrom1C.ToString());
                    _ProjectNumberFrom1C = ProjectFrom1C.ReferenceObject[PM_param_Name_GUID].GetString();
                }
                if (ProjectFrom1C == null)
                {
                    //MessageBox.Show("ProjectFrom1C == null");
                    _ProjectNumberFrom1C = "Не указан";
                }
                return _ProjectNumberFrom1C;
            }
        }

        string _ProjectResponsibleNameFrom1C;
        public string ProjectResponsibleNameFrom1C
        {
            get
            {
                if (_ProjectResponsibleNameFrom1C == null && ProjectFrom1C != null)
                {
                    ReferenceObject responsible = ProjectFrom1C.ReferenceObject.GetObject(PM_link_Responsible_GUID);
                    if (responsible != null)
                        _ProjectResponsibleNameFrom1C = responsible.ToString();
                }
                if (_ProjectResponsibleNameFrom1C == null)
                    _ProjectResponsibleNameFrom1C = "Не указан";
                return _ProjectResponsibleNameFrom1C;
            }

        }

        Guid? _PlanningSpace;
        public Guid PlanningSpace
        {
            get
            {
                if (_PlanningSpace == null)
                {
                    if (Project != null)
                        _PlanningSpace = Project.ReferenceObject[PM_param_PlanningSpace_GUID].GetGuid();
                    else _PlanningSpace = Guid.Empty;
                }
                return (Guid)_PlanningSpace;
            }
        }

        List<ProjectManagementWork> _Children;
        public List<ProjectManagementWork> Children
        {
            get
            {
                if (_Children == null && this.ReferenceObject.Children != null)
                {
                    _Children = new List<ProjectManagementWork>();
                    foreach (var child in this.ReferenceObject.Children)
                    {
                        _Children.Add(Factory.Create_ProjectManagementWork(child));
                    }
                }
                return _Children;
            }
        }

        ProjectManagementWork _ParentWork;
        public ProjectManagementWork ParentWork
        {
            get
            {
                if (_ParentWork == null && this.ReferenceObject.Parent != null)
                    if (this.ReferenceObject.Parent.Class.IsInherit(PM_class_ProjectElement_Guid))
                        _ParentWork = new ProjectManagementWork/* Factory.Create_ProjectManagementWork*/(this.ReferenceObject.Parent);
                return _ParentWork;
            }
        }

        bool? _IsProjectFrom1C;
        public bool IsProjectFrom1C
        {
            get
            {

                if (_IsProjectFrom1C == null)
                {
                    _IsProjectFrom1C = this.ReferenceObject.Class.IsInherit(PM_class_ProjectFrom1C_Guid);
                }
                return (bool)_IsProjectFrom1C;
            }
        }

        /// <summary>
        /// true если ReferenceObject не существует, в противном случае false
        /// </summary>
        public bool IsEmpty
        {
            get
            {
                return ReferenceObject == null;
            }
        }

        bool? _IsProject;
        public bool IsProject
        {
            get
            {

                if (_IsProject == null)
                {
                    _IsProject = this.ReferenceObject.Class.IsInherit(PM_class_Project_Guid);
                }
                return (bool)_IsProject;
            }
        }

        ProjectManagementWork GetTheProject(ProjectManagementWork work)
        {
            //MessageBox.Show("GetTheProject\n" + work.ToString() + "\n" + work.Class.Guid.ToString() + "\n" + PM_class_Project_Guid.ToString());
            if (work.IsProject) return work;
            if (work.ParentWork == null)
            {
                //MessageBox.Show("work.Parent == null");
                return work;
            }
            //MessageBox.Show("work.Parent == " + work.Parent.ToString());
            return GetTheProject(work.ParentWork);
        }
        public ServerConnection Connection { get { return ServerGateway.Connection; } }

        User _Answerable;
        public User Answerable
        {
            get
            {
                if (_Answerable == null)
                {
                    _Answerable = this.ReferenceObject.GetObject(PM_link_Responsible_GUID) as User;
                }
                return _Answerable;
            }
        }

        ProjectManagementWork _ProjectFrom1C;

        public ProjectManagementWork ProjectFrom1C
        {
            get
            {
                if (_ProjectFrom1C == null)
                { _ProjectFrom1C = GetTheProjectFrom1C(this); }
                return _ProjectFrom1C;
            }
        }

        ProjectManagementWork _Project;

        internal ProjectManagementWork Project
        {
            get
            {
                if (_Project == null)
                { _Project = GetTheProject(this); }
                return _Project;
            }
        }

        string _ProjectName;
        public string ProjectName
        {
            get
            {
                if (_ProjectName == null)
                    _ProjectName = this.Project.ToString();
                if (_ProjectName == null)
                    _ProjectName = "Нет проекта";
                return _ProjectName;
            }
        }

        /*List<ReferenceObject> _PlannedUsedResources;
        public IEnumerable<ReferenceObject> PlannedUsedResources
        {
            get
            {
                if (_PlannedUsedResources == null)
                {
                    var usedRes = ReferenceObject.GetObjects(PM_link_UsedResources_GUID).Where(res => (res as ResourcesUsedReferenceObject).ResourcesLink is NonExpendableResourcesReferenceObject);
                    if (usedRes != null)
                    {
                        _PlannedUsedResources = usedRes.Distinct().ToList();
                    }
                    else _PlannedUsedResources = new List<ReferenceObject>();
                    //MessageBox.Show(usedRes.Count().ToString());
                }
                //MessageBox.Show(_PlannedUsedResources.Count().ToString());
                return _PlannedUsedResources;
            }
        }
        */
        double? _Progress;
        public int Progress
        {
            get
            {
                if (_Progress == null)
                {
                    _Progress = ReferenceObject[PM_param_Percent_GUID].GetDouble() * 100;
                }
                return (int)_Progress;
            }
        }

        public DateTime LastEditDate { get; internal set; }

        public override string ToString()
        {
            return this.Name;
        }

        /// <summary>
        /// Изменяет даты начала и окончания на указанные
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        internal void ChangeDates(DateTime startDate, DateTime endDate)
        {
            BeginChanges();
            //if (this.EndDate <= startDate)
            //{
            //    this.ReferenceObject[PM_param_PlanEndDate_GUID].Value = endDate;
            //    this.ReferenceObject[PM_param_PlanStartDate_GUID].Value = startDate;
            //}
            //else// if (this.StartDate >= endDate)
            //{
            this.ReferenceObject[PM_param_PlanStartDate_GUID].Value = startDate;
            this.ReferenceObject.ApplyChanges();
            this.ReferenceObject[PM_param_PlanEndDate_GUID].Value = endDate;
            //}

            this.ReferenceObject.EndChanges(); //если нет изменений то возвращает false

            this.EndDate = endDate;
            this.StartDate = startDate;
        }

        /// <summary>
        /// Изменяет наименование работы
        /// </summary>
        /// <param name="name"></param>
        internal void ChangeWorkName(string name)
        {
            BeginChanges();
            this.ReferenceObject[PM_param_Name_GUID].Value = name;
            this.ReferenceObject.EndChanges(); //если нет изменений то возвращает false
            this._Name = name;
        }


        /// <summary>
        /// получить пользователей которые имеют доступ на элемент проекта "без ограничений"
        /// </summary>
        /// <param name="projectElement"></param>
        /// <param name="accessCode"></param>
        /// <returns></returns>
        public List<User> GetUsersHaveAccess(AccessCode accessCode)
        {
            /// <summary>
            ///  Guid типа "Доступы"
            ///  тип "Список объектов"
            /// </summary>
            Guid TypeAccesses_AccessesList_Guid = new Guid("68989495-719e-4bf3-ba7c-d244194890d5");

            Guid Accesses_TypeParameters_Access_Guid = new Guid("aff97f60-58d6-4c69-807b-dd7b6abfd770"); //Доступ Целое число 

            /// <summary>
            /// Guid параметра - "Уникальный идентификатор пользователя"
            /// Уникальный идентификатор
            /// </summary>
            Guid Accesses_TypeParameters_UserGuid_Guid = new Guid("f65f1d07-275e-44cd-9f9e-5492971e7ebf");
            List<User> users = new List<User>();
            foreach (ReferenceObject access in this.ReferenceObject.GetObjects(TypeAccesses_AccessesList_Guid))//перебор объектов
            {

                if (access[Accesses_TypeParameters_Access_Guid].GetInt32() == (int)accessCode
                    && access[Accesses_TypeParameters_UserGuid_Guid].GetGuid() != Guid.Empty)//Параметры
                {
                    User user = References.UsersReference.Find(access[Accesses_TypeParameters_UserGuid_Guid].GetGuid()) as User;

                    if (!users.Contains(user))
                        users.Add(user);
                }
            }

            return users;
        }


        /// <summary>
        /// переводит объект в состояние редактирования
        /// </summary>
        private void BeginChanges()
        {
            this.ReferenceObject.Reload();//без этой строки вылетает ошибка при взятии на редактирование незаблокированного объекта
            if (this.ReferenceObject == null) this.ReferenceObject = References.WorkChangeRequestReference.CreateReferenceObject(References.Class_WorkChangeRequest);
            else if (this.ReferenceObject.LockState == ReferenceObjectLockState.LockedByCurrentUser)
            {
                this.ReferenceObject.BeginChanges();
            }
            else if (this.ReferenceObject.LockState == ReferenceObjectLockState.LockedByOtherUser)
            {
                this.ReferenceObject.Unlock();
                this.ReferenceObject.BeginChanges();
            }
            else if (this.ReferenceObject.LockState == ReferenceObjectLockState.None)
            {
                this.ReferenceObject.BeginChanges();
            }
            if (!this.ReferenceObject.CanEdit) throw new Exception("Редактирование работы не возможно.");
        }

        internal void Delete()
        {
            this.ReferenceObject.Delete();
            this.ReferenceObject = null;
        }

        internal void AddResource(UsedResource newUsedResource)
        {
            this.BeginChanges();
            newUsedResource.ReferenceObject.Reload();
            this.ReferenceObject.AddLinkedObject(ProjectManagementWork.PM_link_UsedResources_GUID, newUsedResource.ReferenceObject);
            this.ReferenceObject.EndChanges();
            _PlannedNonConsumableResources = null;
        }

        public enum AccessCode
        {
            OnlyViewing,
            NoLimit,
            Parameters,
            Resources = 4,
            ParametersAndResources = 6,
            Detalisation = 8,
            ParametersAndDetalisation = 10,
            ResourcesAndDetalisation = 12,
            ChildWorks = 16,
            ResourcesAndChildWorks = 20,
            DetalisationAndChildWorks = 24
        }

        internal static List<Guid> listParametersToSkip = new List<Guid>() {
        PM_param_PlanningSpace_GUID,//обязательные к пропуску параметры
        PM_param_WorkIntervalLengthDays_Guid, PM_param_WorkIntervalLengthHours_Guid, PM_param_FactStartDate_GUID, PM_param_FactEndDate_GUID,
        PM_param_ProjectCodeFrom1C_Guid};

        internal static List<Guid> listLinksToSkip = new List<Guid>() { PM_list_Access_GUID };

        public static Guid PM_class_ProjectElement_Guid = new Guid("7c968a5b-d1a4-468f-a8cc-d9dbf7ecc990"); //тип - "Элемент проекта"
        public static Guid PM_class_Work_Guid = new Guid("c0bef497-cf64-44a7-9839-a704dc3facb2"); //тип - "Работа"

        static Guid PM_class_ProjectFrom1C_Guid = new Guid("859f7412-c95c-4636-b1f0-f792392ccf83"); //тип - "Затраты проекта 1С new"
        static Guid PM_class_Project_Guid = new Guid("9906a1e9-dc3b-4b27-a083-b3dba9ee8ed0"); // тип - "Проект"
        static Guid PM_list_Access_GUID = new Guid("68989495-719e-4bf3-ba7c-d244194890d5");       //Guid списка "Доступы", справочника - "Управление проектами"
        static Guid PM_param_Name_GUID = new Guid("b7e3b3fe-65c0-4f1f-82e5-3ee95fe360dd");//Guid параметра "Наименование" справочника - "Управление проектами"
        static Guid PM_param_Cipher_GUID = new Guid("b4f4b4d3-5ffc-4869-a880-7561e0c2a574");//Guid параметра "Обозначение" справочника - "Управление проектами"
        static Guid PM_param_PlanningSpace_GUID = new Guid("fd123c69-7945-487a-a854-884ef68f1036");//Guid параметра "Пространство планирования" справочника - "Управление проектами"
        static Guid PM_param_PlanStartDate_GUID = new Guid("4fd8d58a-04ba-4e80-a240-12691e460b83");//Guid параметра "Начало"
        static Guid PM_param_PlanEndDate_GUID = new Guid("da997729-07e2-40aa-a1a0-c08bc4dde481");//Guid параметра "Окончание"
        static Guid PM_param_FactStartDate_GUID = new Guid("2f457df1-246d-4d23-a332-53f387940ba9");      //Guid параметра "Фактическое начало", справочника - "Управление проектами"
        static Guid PM_param_FactEndDate_GUID = new Guid("a93f9644-3e24-4a59-9eb9-d200c83044ee");//Guid параметра "Фактическое окончание"
        static Guid PM_param_UsedResourceIsFact_GUID = new Guid("4227e515-66a4-43ae-9418-346854748986");//Guid параметра "Фактическое значение"
        static Guid PM_param_Percent_GUID = new Guid("f1c3091f-2fff-48d7-a962-cf8d6b08bdd2");//Guid параметра "Процент выполнения"
        static Guid PM_param_ResultOfWork_GUID = new Guid("a24fe48c-cdb4-428f-9330-72fef743b713");//Guid параметра "Результат работы"
        static Guid PM_param_SumPlanResourceCount_GUID = new Guid("cde1942c-ffc6-4654-9a2a-8f0b3a0a8e14");//Guid параметра "Суммарные плановые трудозатраты по ресурсам"
        static Guid PM_param_ProjectCodeFrom1C_Guid = new Guid("b37cfad2-bd50-47f3-8292-8f7c10d8b9a2"); // Guid параметра "Код проекта из 1С", справочника - "Управление проектами"
        static Guid PM_param_WorkIntervalLengthDays_Guid = new Guid("488646b9-a912-468f-a791-e9e68774215f");      //Guid параметра "Длительность(Дни)", справочника - "Управление проектами"
        static Guid PM_param_WorkIntervalLengthHours_Guid = new Guid("9f733670-5f3b-4554-b023-e4d450ab0c00");      //Guid параметра "Длительность", справочника - "Управление проектами"

        public static readonly Guid PM_link_UsedResources_GUID = new Guid("1a22ee46-5438-4caa-8b75-8a8a37b74b7e");//Guid списка 1:n "Ресурс"
        static Guid PM_link_Responsible_GUID = new Guid("063df6fa-3889-4300-8c7a-3ce8408a931a");//Guid связи N:1 "Ответственный"

    }

    public interface ITreeNode
    {
        ITreeNode Parent { get; }

        IEnumerable<ITreeNode> Children { get; }

    }

    public class ProjectTreeItem : ITreeNode
    {
        //public ProjectTreeItem(ReferenceObject referenceObject)
        //{
        //    this.ReferenceObject = referenceObject;
        //}
        public ReferenceObject ReferenceObject;

        ITreeNode _Parent;
        public ITreeNode Parent
        {
            get
            {
                if (_Parent == null)
                {
                    ReferenceObject parent = this.ReferenceObject.Parent;
                    if (parent != null) _Parent = Factory.CreateProjectTreeItem(parent);
                }
                return _Parent;
            }
        }

        IEnumerable<ITreeNode> _Children;
        public IEnumerable<ITreeNode> Children
        {
            get
            {
                if (_Children == null)
                {
                    List<ITreeNode> temp = new List<ITreeNode>();
                    foreach (var item in this.ReferenceObject.Children)
                    {
                        ITreeNode node = Factory.CreateProjectTreeItem(item);
                        if (node != null) temp.Add(node);
                    }
                    _Children = temp;
                }
                return _Children;
            }
        }
    }

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

    static class Factory
    {

        static void FillCreationTable()
        {
            if (CreationTable == null)
            {
                CreationTable = new Dictionary<ClassObject, Type>();
                CreationTable.Add(References.Class_WorkChangeRequest_DatesChange, typeof(WorkChangeRequest_DatesChange));
                CreationTable.Add(References.Class_WorkChangeRequest_WorkNameChange, typeof(WorkChangeRequest_WorkNameChange));
                CreationTable.Add(References.Class_WorkChangeRequest_ChangeResource, typeof(WorkChangeRequest_ChangeResource));
                CreationTable.Add(References.Class_WorkChangeRequest_DeleteWork, typeof(WorkChangeRequest_DeleteWork));
                CreationTable.Add(References.Class_WorkChangeRequest_AddWork, typeof(WorkChangeRequest_AddWork));
            }
        }

        static Dictionary<ClassObject, Type> CreationTable;
        public static dynamic Create_WorkChangeRequest(ReferenceObject referenceObject)
        {
            //var officialNoteBase = TryGetFromCache(referenceObject);
            //if (officialNoteBase != null) return officialNoteBase;
            // MessageBox.Show("CreationTable.count = " + CreationTable.Count.ToString());
            //Guid param_RequestType_Guid = new Guid("3027c606-32a3-4c2f-b7e5-2a2c5ebc6817"); //Guid параметра "Тип запроса" (Длина 64 символов) Field_6001
            //string typeName = referenceObject[param_RequestType_Guid].GetString();
            ClassObject className = referenceObject.Class;
            System.Type type;
            if (CreationTable.TryGetValue(className, out type))
            {
                return Activator.CreateInstance(type, referenceObject);
            }
            else throw new ArgumentOutOfRangeException("Не представимый тип входящего аргумента!" + string.Format(". Reference - {0}, ObjectGuid - {1}", referenceObject.Reference.Name, referenceObject.SystemFields.Guid.ToString()));
        }

        public static dynamic Create_WorkChangeRequest(EditReasonNote note)
        {
            //var officialNoteBase = TryGetFromCache(referenceObject);
            //if (officialNoteBase != null) return officialNoteBase;
            //MessageBox.Show("CreateNew" + List_OfficialNotes_Cache.Count.ToString());
            //Guid param_RequestType_Guid = new Guid("3027c606-32a3-4c2f-b7e5-2a2c5ebc6817"); //Guid параметра "Тип запроса" (Длина 64 символов) Field_6001
            string typeName = note.TypeName;
            return Create_WorkChangeRequest_ByTypeName(typeName, note);
        }

        static dynamic Create_WorkChangeRequest_ByTypeName(string typeName, dynamic note)
        {
            if (typeName == WorkChangeRequest.TypeName.DatesCorrection)
            {
                var result = new WorkChangeRequest_DatesChange(note);
                return result;
            }
            else if (typeName == WorkChangeRequest.TypeName.AddChildWork)
            {
                var result = new WorkChangeRequest_AddWork(note);
                return result;
            }
            else if (typeName == WorkChangeRequest.TypeName.ChangeResource)
            {
                var result = new WorkChangeRequest_ChangeResource(note);
                return result;
            }
            else if (typeName == WorkChangeRequest.TypeName.ChangeWorkName)
            {
                var result = new WorkChangeRequest_WorkNameChange(note);
                return result;
            }
            else if (typeName == WorkChangeRequest.TypeName.DeleteChildWork)
            {
                var result = new WorkChangeRequest_DeleteWork(note);
                return result;
            }
            throw new Exception("Невозможно создать запрос на основе данной записи журнала редактирования, данный тип не поддерживается!\n" + typeName);
        }

        public static dynamic Create_ProjectManagementWork(ReferenceObject referenceObject)
        {
            ProjectManagementWork projectManagementWork = TryGetFromCache_ProjectManagementWork(referenceObject);
            if (projectManagementWork == null)
            {
                projectManagementWork = new ProjectManagementWork(referenceObject);
                List_ProjectManagementWorks_Cache.Add(projectManagementWork);
            }
            return projectManagementWork;
        }

        private static ProjectManagementWork TryGetFromCache_ProjectManagementWork(ReferenceObject referenceObject)
        {
            List<ProjectManagementWork> listFoundedObject = new List<ProjectManagementWork>();
            foreach (var item in List_ProjectManagementWorks_Cache)
            {
                if (item.ReferenceObject.SystemFields.Id == referenceObject.SystemFields.Id) listFoundedObject.Add(item);
            }
            //var listFoundedDobject = List_ProjectManagementWorks_Cache.Where(on => on.ReferenceObject.SystemFields.Id == referenceObject.SystemFields.Id).ToList();
            if (listFoundedObject == null) return null;
            ProjectManagementWork currentVersion = null;
            foreach (var item in listFoundedObject)
            {
                if (referenceObject.SystemFields.EditDate != item.LastEditDate)
                {
                    List_ProjectManagementWorks_Cache.Remove(item);
                }
                else currentVersion = item;
            }
            return currentVersion;
        }

        internal static ITreeNode CreateProjectTreeItem(ReferenceObject referenceObject)
        {
            //если это работа то создаём работу
            if (referenceObject.Class.IsInherit(References.ProjectManagementReference.Classes.Find(ProjectManagementWork.PM_class_ProjectElement_Guid))) return Create_ProjectManagementWork(referenceObject);
            if (referenceObject.Class.IsInherit(References.ProjectManagementReference.Classes.Find(ProjectFolder.PM_class_ProjectFolder_Guid))) return new ProjectFolder(referenceObject);
            return null;
        }

        //private static OfficialNoteBase TryGetFromCache(ReferenceObject referenceObject)
        //{
        //    return null;
        //    var listFoundedObjects = List_OfficialNotes_Cache.Where(on => on.ReferenceObject.SystemFields.Id == referenceObject.SystemFields.Id).ToList();
        //    if (listFoundedObjects == null) return null;
        //    OfficialNoteBase currentVersion = null;
        //    foreach (var item in listFoundedObjects)
        //    {
        //        if (referenceObject.SystemFields.EditDate != item.LastEditDate)
        //        {
        //            List_OfficialNotes_Cache.Remove(item);
        //        }
        //        else currentVersion = item;
        //    }
        //    return currentVersion;
        //}

        static Factory()
        {
            FillCreationTable();
            if (List_ProjectManagementWorks_Cache == null) List_ProjectManagementWorks_Cache = new List<ProjectManagementWork>();
            //if (List_OfficialNotes_Cache == null) List_OfficialNotes_Cache = new List<OfficialNoteBase>();
        }

        static List<ProjectManagementWork> List_ProjectManagementWorks_Cache;
        //static List<OfficialNoteBase> List_OfficialNotes_Cache;
    }

    /// <summary>
    /// Для работы с синхронизацией в УП
    /// ver 1.3
    /// </summary>
    class Synchronization
    {
        /// <summary>
        /// Возвращает список синхронизированных работ из указанного пространства планирования
        /// </summary>
        /// <param name="work"></param>
        /// <param name="planningSpaceGuidString"></param>
        /// <param name="returnOnlyMasterWorks"> если true только укрупнения, если false только детализации</param>
        /// <returns></returns>
        public static List<ProjectManagementWork> GetSynchronizedWorksInProject(ProjectManagementWork work, ProjectManagementWork project, bool? returnOnlyMasterWorks = null)
        {
            string guidWorkForSearch = work.Guid.ToString();
            string guidProjectForSearch = project.Guid.ToString();
            Filter filter = new Filter(ProjectDependenciesReferenceInfo);
            // Условие: тип зависимости "синхронизация"
            ReferenceObjectTerm termSyncType = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
            // устанавливаем параметр
            termSyncType.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_DependencyType_Guid));
            // устанавливаем оператор сравнения
            termSyncType.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termSyncType.Value = 5;// DependencyType.Synchronization;

            // Условие: в названии содержится "является укрупнением"
            //ReferenceObjectTerm termMasterWork = new ReferenceObjectTerm(LogicalOperator.Or);

            ReferenceObjectTerm termMasterWork = new ReferenceObjectTerm(filter.Terms);
            // устанавливаем параметр
            termMasterWork.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_MasterWork_Guid));
            // устанавливаем оператор сравнения
            termMasterWork.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termMasterWork.Value = guidWorkForSearch;

            ReferenceObjectTerm termSlaveProjectWork = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
            // устанавливаем параметр
            termSlaveProjectWork.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_SlaveProject_Guid));
            // устанавливаем оператор сравнения
            termSlaveProjectWork.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termSlaveProjectWork.Value = guidProjectForSearch;

            // Условие: в названии содержится "является детализацией"
            ReferenceObjectTerm termSlaveWork = new ReferenceObjectTerm(filter.Terms, LogicalOperator.Or);
            //ReferenceObjectTerm term1 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.Or);
            // устанавливаем параметр
            termSlaveWork.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_SlaveWork_Guid));
            // устанавливаем оператор сравнения
            termSlaveWork.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termSlaveWork.Value = guidWorkForSearch;

            ReferenceObjectTerm termMasterProjectWork = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
            // устанавливаем параметр
            termMasterProjectWork.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_MasterProject_Guid));
            // устанавливаем оператор сравнения
            termMasterProjectWork.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termMasterProjectWork.Value = guidProjectForSearch;

            // Группируем условия в отдельную группу (другими словами добавляем скобки)
            TermGroup group1 = filter.Terms.GroupTerms(new Term[] { termMasterWork, termSlaveProjectWork });
            TermGroup group2 = filter.Terms.GroupTerms(new Term[] { termSlaveWork, termMasterProjectWork });

            //редактируем при необходимости
            if (returnOnlyMasterWorks != null)
            {
                if ((bool)returnOnlyMasterWorks) group1.Clear();
                else group2.Clear();
            }

            //Получаем список объектов, в качестве условия поиска – сформированный фильтр
            List<ReferenceObject> listObj = ProjectDependenciesReference.Find(filter);
            List<Guid> listSyncWorkGuids = new List<Guid>();
            //#if test
            //            System.Windows.Forms.MessageBox.Show(filter.ToString() + "\nlistObj.Count = " + listObj.Count.ToString());
            //#endif
            foreach (var item in listObj) // выбираем все отличные от исходной работы guids
            {
                Guid slaveGuid = item[SynchronizationParameterGuids.param_SlaveWork_Guid].GetGuid();
                Guid masterGuid = item[SynchronizationParameterGuids.param_MasterWork_Guid].GetGuid();
                if (slaveGuid != work.Guid) listSyncWorkGuids.Add(slaveGuid);
                if (masterGuid != work.Guid) listSyncWorkGuids.Add(masterGuid);
            }
            listSyncWorkGuids = listSyncWorkGuids.Distinct().ToList();
            List<ProjectManagementWork> result = new List<ProjectManagementWork>();

            foreach (var item in listSyncWorkGuids)
            {
                ProjectManagementWork tempWork = new ProjectManagementWork(item);
                result.Add(new ProjectManagementWork(item));

            }
            return result;


        }

        [Obsolete]
        /// <summary>
        /// Use GetSyncronizedWorks
        /// Возвращает список синхронизированных работ из указанного пространства планирования
        /// </summary>
        /// <param name="work"></param>
        /// <param name="planningSpaceGuidString"></param>
        /// <param name="returnOnlyMasterWorks"> если true только укрупнения, если false только детализации</param>
        /// <returns></returns>
        public static List<ProjectManagementWork> GetSynchronizedWorksFromSpace(ProjectManagementWork work, string planningSpaceGuidString, bool? returnOnlyMasterWorks = null)
        {
            string guidStringForSearch = work.Guid.ToString();
            Filter filter = GetFilter(returnOnlyMasterWorks, guidStringForSearch);

            //Получаем список объектов, в качестве условия поиска – сформированный фильтр
            List<ReferenceObject> listObj = ProjectDependenciesReference.Find(filter);
            List<Guid> listSyncWorkGuids = new List<Guid>();
            //#if test
            //            MessageBox.Show(filter.ToString() + "\nlistObj.Count = " + listObj.Count.ToString());
            //#endif
            foreach (var item in listObj) // выбираем все отличные от исходной работы guids
            {
                Guid slaveGuid = item[SynchronizationParameterGuids.param_SlaveWork_Guid].GetGuid();
                Guid masterGuid = item[SynchronizationParameterGuids.param_MasterWork_Guid].GetGuid();
                if (slaveGuid != work.Guid) listSyncWorkGuids.Add(slaveGuid);
                if (masterGuid != work.Guid) listSyncWorkGuids.Add(masterGuid);
            }
            listSyncWorkGuids = listSyncWorkGuids.Distinct().ToList();
            List<ProjectManagementWork> result = new List<ProjectManagementWork>();
            foreach (var item in listSyncWorkGuids)
            {
                ProjectManagementWork tempWork = new ProjectManagementWork(item);
                if (tempWork.PlanningSpace.ToString() == planningSpaceGuidString)
                {
                    result.Add(new ProjectManagementWork(item));
                }
            }
            return result;


        }

        /// <summary>
        /// Возвращает список синхронизированных работ из указанного пространства планирования
        /// </summary>
        /// <param name="work"></param>
        /// <param name="planningSpaceGuidString"></param>
        /// <param name="returnOnlyMasterWorks"> если true только укрупнения, если false только детализации</param>
        /// <returns></returns>
        public static List<ProjectManagementWork> GetSynchronizedWorks(ProjectManagementWork work, string planningSpaceGuidString, bool? returnOnlyMasterWorks)
        {
            string guidStringForSearch = work.Guid.ToString();
            Filter filter = GetFilter(returnOnlyMasterWorks, guidStringForSearch);

            //Получаем список объектов, в качестве условия поиска – сформированный фильтр
            List<ReferenceObject> listObj = ProjectDependenciesReference.Find(filter);
            List<Guid> listSyncWorkGuids = new List<Guid>();
            //#if test
            //            MessageBox.Show(filter.ToString() + "\nlistObj.Count = " + listObj.Count.ToString());
            //#endif
            foreach (var item in listObj) // выбираем все отличные от исходной работы guids
            {
                Guid slaveGuid = item[SynchronizationParameterGuids.param_SlaveWork_Guid].GetGuid();
                Guid masterGuid = item[SynchronizationParameterGuids.param_MasterWork_Guid].GetGuid();
                if (slaveGuid != work.Guid) listSyncWorkGuids.Add(slaveGuid);
                if (masterGuid != work.Guid) listSyncWorkGuids.Add(masterGuid);
            }
            listSyncWorkGuids = listSyncWorkGuids.Distinct().ToList();
            List<ProjectManagementWork> result = new List<ProjectManagementWork>();
            bool isNeedToFilterByPlanningSpace = !string.IsNullOrWhiteSpace(planningSpaceGuidString);
            foreach (var item in listSyncWorkGuids)
            {
                ProjectManagementWork tempWork = new ProjectManagementWork(item);
                if (isNeedToFilterByPlanningSpace && tempWork.PlanningSpace.ToString() != planningSpaceGuidString)
                {
                    continue;
                }
                result.Add(new ProjectManagementWork(item));

            }
            return result;

        }

        private static Filter GetFilter(bool? returnOnlyMasterWorks, string guidStringForSearch)
        {
            Filter filter = new Filter(ProjectDependenciesReferenceInfo);
            // Условие: тип зависимости "синхронизация"
            ReferenceObjectTerm termSyncType = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
            // устанавливаем параметр
            termSyncType.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_DependencyType_Guid));
            // устанавливаем оператор сравнения
            termSyncType.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termSyncType.Value = 5;

            // Условие: в названии содержится "является укрупнением"
            //ReferenceObjectTerm termMasterWork = new ReferenceObjectTerm(LogicalOperator.Or);

            ReferenceObjectTerm termMasterWork = new ReferenceObjectTerm(filter.Terms);
            // устанавливаем параметр
            termMasterWork.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_MasterWork_Guid));
            // устанавливаем оператор сравнения
            termMasterWork.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termMasterWork.Value = guidStringForSearch;
            // Условие: в названии содержится "является детализацией"
            ReferenceObjectTerm termSlaveWork = new ReferenceObjectTerm(filter.Terms, LogicalOperator.Or);
            //ReferenceObjectTerm term1 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.Or);
            // устанавливаем параметр
            termSlaveWork.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_SlaveWork_Guid));
            // устанавливаем оператор сравнения
            termSlaveWork.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termSlaveWork.Value = guidStringForSearch;

            // Группируем условия в отдельную группу (другими словами добавляем скобки)
            TermGroup group1 = filter.Terms.GroupTerms(new Term[] { termMasterWork, termSlaveWork });

            //редактируем при необходимости
            if (returnOnlyMasterWorks != null)
            {
                if (!(bool)returnOnlyMasterWorks) group1.Remove(termSlaveWork);
                else group1.Remove(termMasterWork);
            }

            return filter;
        }

        //создаём синхронизацию
        public static void SyncronizeWorks(ProjectManagementWork masterWork, ProjectManagementWork slaveWork)
        {
            ReferenceObject синхронизация = ProjectDependenciesReference.CreateReferenceObject(Class_ProjectDependency);
            синхронизация[SynchronizationParameterGuids.param_DependencyType_Guid].Value = 5;// DependencyType.Synchronization;
            синхронизация[SynchronizationParameterGuids.param_MasterWork_Guid].Value = masterWork.Guid;
            синхронизация[SynchronizationParameterGuids.param_SlaveWork_Guid].Value = slaveWork.Guid;
            синхронизация[SynchronizationParameterGuids.param_MasterProject_Guid].Value = masterWork.Project.Guid;
            синхронизация[SynchronizationParameterGuids.param_SlaveProject_Guid].Value = slaveWork.Project.Guid;
            синхронизация.EndChanges();
            //return синхронизация;
        }

        /// <summary>
        /// Удаляет синхронизацию между указанными работами
        /// </summary>
        /// <param name="masterWork"></param>
        /// <param name="slaveWork"></param>
        public static void DeleteSynchronisationBetween(ProjectManagementWork masterWork, ProjectManagementWork slaveWork)
        {

            Filter filter = new Filter(ProjectDependenciesReferenceInfo);
            // Условие: тип зависимости "синхронизация"
            ReferenceObjectTerm termSyncType = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
            // устанавливаем параметр
            termSyncType.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_DependencyType_Guid));
            // устанавливаем оператор сравнения
            termSyncType.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termSyncType.Value = 5;

            // Условие: в названии содержится "является укрупнением"
            string masterWorkGuidStringForSearch = masterWork.Guid.ToString();

            ReferenceObjectTerm termMasterWork = new ReferenceObjectTerm(filter.Terms);
            // устанавливаем параметр
            termMasterWork.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_MasterWork_Guid));
            // устанавливаем оператор сравнения
            termMasterWork.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termMasterWork.Value = masterWorkGuidStringForSearch;

            // Условие: в названии содержится "является детализацией"
            string slaveWorkGuidStringForSearch = slaveWork.Guid.ToString();
            ReferenceObjectTerm termSlaveWork = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
            //ReferenceObjectTerm term1 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.Or);
            // устанавливаем параметр
            termSlaveWork.Path.AddParameter(ProjectDependenciesReference.ParameterGroup.OneToOneParameters.Find(SynchronizationParameterGuids.param_SlaveWork_Guid));
            // устанавливаем оператор сравнения
            termSlaveWork.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termSlaveWork.Value = slaveWorkGuidStringForSearch;


            //Получаем список объектов, в качестве условия поиска – сформированный фильтр
            List<ReferenceObject> listObj = ProjectDependenciesReference.Find(filter);
            //#if test
            //            MessageBox.Show(filter.ToString() + "\nlistObj.Count = " + listObj.Count.ToString());
            //#endif
            foreach (var item in listObj) // удаляем всё что нашли
            {
                item.Delete();
            }
        }


        private static ClassObject _Class_ProjectDependency;
        /// <summary>
        /// класс Запрос коррекции работ
        /// </summary>
        public static ClassObject Class_ProjectDependency
        {
            get
            {
                if (_Class_ProjectDependency == null)
                    _Class_ProjectDependency = ProjectDependenciesReference.Classes.Find(SynchronizationParameterGuids.class_Dependency_GUID);

                return _Class_ProjectDependency;
            }
        }


        static Reference _ProjectDependenciesReference;
        /// <summary>
        /// Справочник "Зависимости проектов"
        /// </summary>
        public static Reference ProjectDependenciesReference
        {
            get
            {
                if (_ProjectDependenciesReference == null)
                    _ProjectDependenciesReference = GetReference(ref _ProjectDependenciesReference, ProjectDependenciesReferenceInfo);

                return _ProjectDependenciesReference;
            }
        }

        static ReferenceInfo _ProjectDependenciesReferenceInfo;
        /// <summary>
        /// Справочник "Зависимости проектов"
        /// </summary>
        private static ReferenceInfo ProjectDependenciesReferenceInfo
        {
            get { return GetReferenceInfo(ref _ProjectDependenciesReferenceInfo, SynchronizationParameterGuids.ref_ProjDependencies_Guid); }
        }

        private static Reference GetReference(ref Reference reference, ReferenceInfo referenceInfo)
        {
            if (reference == null)
                reference = referenceInfo.CreateReference();


            return reference;
        }
        private static ReferenceInfo GetReferenceInfo(ref ReferenceInfo referenceInfo, Guid referenceGuid)
        {
            if (referenceInfo == null)
                referenceInfo = References.Connection.ReferenceCatalog.Find(referenceGuid);

            return referenceInfo;
        }
        struct SynchronizationParameterGuids
        {
            //Справочник - "Зависимости проектов"
            public static readonly Guid ref_ProjDependencies_Guid = new Guid("e13cee45-39fa-43ff-ba2a-957294d975bf");    //Guid справочника - "Зависимости проектов"
            public static readonly Guid class_Dependency_GUID = new Guid("0ed0174e-5392-460f-8b53-5a2e52c26f9b");       //Guid типа "Зависимости проектов" справочника - "Зависимости проектов"
            public static readonly Guid param_MasterProject_Guid = new Guid("087648ca-6269-4e7f-8e5a-9fabbab5fafd");    //Guid параметра "Проект 1"
            public static readonly Guid param_MasterWork_Guid = new Guid("17feb793-7656-475a-b7b0-61672ba813fa");    //Guid параметра "Работа 1"
            public static readonly Guid param_SlaveProject_Guid = new Guid("a5590629-9469-4c39-87e6-169b60678abb");    //Guid параметра "Проект 2"
            public static readonly Guid param_SlaveWork_Guid = new Guid("5baf8272-f639-4fc8-af65-6f0af32647fd");    //Guid параметра "Работа 2"
            public static readonly Guid param_DependencyType_Guid = new Guid("45bc3bf0-88cf-4b46-b951-b1e07a5d7fcb");    //Guid параметра "Тип зависимости"
        }
    }

    struct PlanningSpaceGuidString
    {
        public const string Departments = "d2b9ed71-5ff9-4916-97e2-c05f06840ffb";
        public const string WorkingPlans = "7357c19d-909b-44f8-aae1-13999fbd7605";
    }

    /// <summary>
    /// HTML письмо
    /// ver 1.6
    /// Закомментирован метод Show()
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
            //MailMessage.
        }

        //#if !server
        //        public void Show()
        //        {
        //            MailMessage = new MailMessage(Connection.Mail.DOCsAccount);
        //            MailMessage.Subject = CleanString(Subject);
        //            string messageBody = CreateMessageBody();
        //            MailMessage.SetBody(messageBody, MailBodyType.Html);
        //            FillAttachments();
        //            FillLocalFiles();
        //            FillMailRecipients();
        //            FillMailCopyRecipients();
        //            TFlex.DOCs.UI.DialogManager.Instance.ShowPropertyDialog(MailMessage);
        //        }
        //#endif
        private void FillLocalFiles()
        {
            if (_pathesToLocalAttachments == null) return;
            foreach (var localFilePath in _pathesToLocalAttachments)
            {
                FileAttachment fileAttachment = new FileAttachment(localFilePath);
                MailMessage.Attachments.Add(fileAttachment);
                //fileAttachment.DownloadFile();

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
<img src=""http://portal.dinamika-avia.ru/include/logo.2186.png""  alt=""АСУ ""Динамика"""">
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
<br>{0}
<br>
<img src=""http://portal.dinamika-avia.ru/include/logo.2186.png""  alt=""АСУ ""Динамика"""">
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

    [System.Serializable]
    public class WorkNotFoundException : Exception
    {
        public WorkNotFoundException() { }
        public WorkNotFoundException(string message) : base(message) { }
        public WorkNotFoundException(string message, Exception inner) : base(message, inner) { }
        protected WorkNotFoundException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }

    [System.Serializable]
    public class DepartmentWorkNotFoundException : WorkNotFoundException
    {
        public DepartmentWorkNotFoundException() { }
        public DepartmentWorkNotFoundException(string message) : base(message) { }
        public DepartmentWorkNotFoundException(string message, Exception inner) : base(message, inner) { }
        protected DepartmentWorkNotFoundException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }

    [System.Serializable]
    public class ManagerWorkNotFoundException : WorkNotFoundException
    {
        public ManagerWorkNotFoundException() { }
        public ManagerWorkNotFoundException(string message) : base(message) { }
        public ManagerWorkNotFoundException(string message, Exception inner) : base(message, inner) { }
        protected ManagerWorkNotFoundException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }
}

