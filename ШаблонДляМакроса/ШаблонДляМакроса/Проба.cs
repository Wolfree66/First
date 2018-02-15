//#define server
//#define test

namespace Cancelaria
{
    using System;
    //using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Windows.Forms;
    using TFlex.DOCs.Model;
    using TFlex.DOCs.Model.Classes;
    using TFlex.DOCs.Model.Macros;
    using TFlex.DOCs.Model.References;
    using TFlex.DOCs.Model.References.Files;
    using TFlex.DOCs.Model.References.Users;
    using TFlex.DOCs.Model.Macros.ObjectModel;
    using TFlex.DOCs.Model.Mail;
    using TFlex.DOCs.Model.Stages;
    using TFlex.DOCs.Model.Structure;
    using TFlex.DOCs.UI.Objects.Managers;
    using TFlex.DOCs.UI.Common.References;
    using TFlex.DOCs.UI.Common;
    using TFlex.DOCs.Model.Search;
    using TFlex.DOCs.UI;


    public class Macro : MacroProvider
    {
        public Macro(MacroContext context)
            : base(context)
        {
        }

        private static Guid reference_Users_Guid = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");    //Guid справочника - "Группы и пользователи"
        private static Guid reference_RCC_Guid = new Guid("80831dc7-9740-437c-b461-8151c9f85e59");    //Guid справочника - "Регистрационно-контрольные карточки"
        private static Guid RCC_link_ExtraAccess_GUID = new Guid("7f823151-f3ca-40be-97b0-a1c69126a027"); //связь со справочником Группы и пользователи - Дополнительный доступ
        private static Guid RCC_link_SignedTo_GUID = new Guid("a3746fbd-7bda-4326-8e4b-f6bbefc2ecce"); //связь N:N со Подписал/Кому
        private static Guid RCC_link_Responsible_GUID = new Guid("6a69454d-7803-4cd8-8874-cde190185b7b"); //связь N:1  Ответственный
        private static Guid RCC_link_Performer_GUID = new Guid("17330c2c-589f-4838-b6d1-8a9214666c4e"); //связь N:1  Исполнитель



        public override void Run()
        {
            MailRecipient mailRecipient = new MailRecipient("Липаев", "lipaev@dinamika-avia.ru");
            ReferenceObject rccOut = References.RegistryControlCards.Find(35018);
            ReferenceObject ro = rccOut.GetObject(RegistryControlCard.RCC_link_RegistryJournal_GUID);
            MessageBox.Show(ro.ToString());
            RegistryControlCard registryControlCard = new RegistryControlCard(rccOut);
            //ОтправитьПисьмоОНовойЗаписиВКанцелярии(registryControlCard, new List<MailRecipient> { mailRecipient });
            MessageBox.Show(registryControlCard.RegistryJournal.ToString());
            MessageBox.Show(registryControlCard.RegistryJournal.Name);
        }

        public void Test()
        {

            //ReferenceObject ro = GetMailShipmentByName("Электронная почта");
            ReferenceObject rccOut = References.RegistryControlCards.Find(35018);

            RegistryControlCard registryControlCard = new RegistryControlCard(rccOut);
            //MessageBox.Show(registryControlCard.GetAllLinkedFiles().Count().ToString());
            //MessageBox.Show(registryControlCard.Documents.Count().ToString());
            //return;
            SendInternalRegistryControlCard(rccOut);

            IEnumerable<Organization> toOrganizations = registryControlCard.Get_ToOrganizations();
            if (!IsOrganizationsEmailsAreCorrect(toOrganizations)) return;

            if (registryControlCard.Performer == null)
            {
                MessageBox.Show("Не указан исполнитель!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            //создаём список адресатов
            List<MailRecipient> mailRecipients = (from org in toOrganizations
                                                  select new MailRecipient(org.Name, org.Email)).ToList();
            List<MailRecipient> mailCopyRecipients = new List<MailRecipient> { new MailRecipient(registryControlCard.Performer as UserReferenceObject) };
            foreach (var item in mailRecipients)
            {
                string na = item.ToString();
            }
            foreach (var item in mailCopyRecipients)
            {
                string na = item.ToString();
            }
            //показываем диалог выбора адресатов
        }

        #region Замена организации во всех записях канцелярии

        public void ЗаменитьОрганизациюУВсехОбъектовКанцелярии()
        {
            //выбрать организации заменяемую и новую
            Organization заменяемаяОрганизация = null;
            Organization новаяОрганизация = null;
            if (!ВыбираемЦелевуюОрганизациюИЗамену(ref заменяемаяОрганизация, ref новаяОрганизация)) return;
            //вывести подтверждение замены
            if (MessageBox.Show(string.Format("Заменить организацию {0} на {1}, во всех записях канцелярии?", заменяемаяОрганизация.ToString(), новаяОрганизация.ToString()), "Подтвердите операцию", MessageBoxButtons.OKCancel) == DialogResult.Cancel) return;

            //найти все РКК где организацией является заменяемая 
            List<RegistryControlCard> listRCC = FindAllRCCWithOrganization(заменяемаяОрганизация);

            int success = 0;
            int fault = 0;
            int all = listRCC.Count;
            StringBuilder errorLog = new StringBuilder();
            int counter = 0;
            WaitingHelper.Wait(null, "Идёт замена", false, new Action<int>((k) =>
            {
                //заменить организацию на новую
                foreach (var item in listRCC)
                {
                    WaitingHelper.SetText(string.Format("Замена организации {0} из {1}", counter, all));
                    counter++;
                    try
                    {
                        ReplaceOrganization(item, заменяемаяОрганизация, новаяОрганизация);
                        success++;
                    }
                    catch//в случае ошибки записать в лог
                    {
                        errorLog.AppendLine(string.Format("{0} №{1};", item.Type, item.RegistryNumber));
                        fault++;
                    }
                }
            }), counter);
            //показать результат замены
            string result = string.Format("Успешно заменено - {0}/{1}; ошибок - {2}", success, all, fault);
            if (errorLog.Length > 0) MessageBox.Show(result + "\n\nВозникли ошибки в следующих записях:\n" + errorLog.ToString(), "Не все записи были обработаны", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
            {
                //заменяемаяОрганизация.Delete();
                MessageBox.Show(result, "Результат операции");
            }
        }


        private void ReplaceOrganization(RegistryControlCard registryCard, Organization заменяемаяОрганизация, Organization новаяОрганизация)
        {
#if !test
            ReferenceObject referenceObjectRCC = registryCard.ReferenceObject;
            referenceObjectRCC.BeginChanges();
            referenceObjectRCC.RemoveLinkedObject(RegistryControlCard.RCC_link_Organizations_GUID, заменяемаяОрганизация.ReferenceObject);
            referenceObjectRCC.AddLinkedObject(RegistryControlCard.RCC_link_Organizations_GUID, новаяОрганизация.ReferenceObject);
            referenceObjectRCC.EndChanges();
#endif
        }

        private List<RegistryControlCard> FindAllRCCWithOrganization(Organization заменяемаяОрганизация)
        {
            Filter filter = new Filter(References.RegistryControlCardsReferenceInfo);
            // Условие: связь "Организации" содержит
            ReferenceObjectTerm termSyncType = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
            // устанавливаем параметр
            ParameterGroup linkToOrganization = References.RegistryControlCards.ParameterGroup.OneToManyLinks.Find(RegistryControlCard.RCC_link_Organizations_GUID);
            termSyncType.Path.AddGroup(linkToOrganization);
            termSyncType.Path.AddParameter(linkToOrganization.SlaveGroup[SystemParameterType.ObjectId]);
            // устанавливаем оператор сравнения
            termSyncType.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            termSyncType.Value = заменяемаяОрганизация.ID;
#if test
            MessageBox.Show(filter.ToString());
#endif
            //Получаем список объектов, в качестве условия поиска – сформированный фильтр
            List<ReferenceObject> listObj = References.RegistryControlCards.Find(filter);
            List<RegistryControlCard> result = new List<RegistryControlCard>();
            foreach (var item in listObj)
            { result.Add(new RegistryControlCard(item)); }
            return result;
        }

        private bool ВыбираемЦелевуюОрганизациюИЗамену(ref Organization заменяемаяОрганизация, ref Organization новаяОрганизация)
        {
            ДиалогВвода диалог = СоздатьДиалогВвода("Выберите заменяемую и новую организации");
            диалог.УстановитьРазмер(600, 0);

            диалог.ДобавитьВыборИзСправочника("Организация которую нужно заменить", "Данные Организаций", null, true);
            //Объект объект = null;//Объект.CreateInstance(целевойПроект, Context);
            диалог.ДобавитьВыборИзСправочника("Замена", "Данные Организаций", null, true);

            //диалог["Выберите целевой проект"] = целевойПроект;
            if (диалог.Показать())
            {
                ReferenceObject заменяемаяОрганизацияReferenceObject = (ReferenceObject)диалог["Организация которую нужно заменить"];
                ReferenceObject новаяОрганизацияReferenceObject = (ReferenceObject)диалог["Замена"];
                //MessageBox.Show(заменяемаяОрганизацияReferenceObject.ToString());
                //MessageBox.Show(новаяОрганизацияReferenceObject.ToString());
                if (заменяемаяОрганизацияReferenceObject != null && новаяОрганизацияReferenceObject != null)
                {
                    if (заменяемаяОрганизацияReferenceObject != новаяОрганизацияReferenceObject)
                    {
                        заменяемаяОрганизация = new Organization(заменяемаяОрганизацияReferenceObject);
                        новаяОрганизация = new Organization(новаяОрганизацияReferenceObject);
                        return true;
                    }
                    else
                    {
                        MessageBox.Show("Выбрана одна и та же организация!\nОрганизации должны быть разными.", "Выберите разные организации!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return ВыбираемЦелевуюОрганизациюИЗамену(ref заменяемаяОрганизация, ref новаяОрганизация);
                    }
                }
                else
                {
                    MessageBox.Show("Необходимо выбрать организацию!", "Необходимо выбрать организацию!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return ВыбираемЦелевуюОрганизациюИЗамену(ref заменяемаяОрганизация, ref новаяОрганизация);
                }
            }

            return false;
        }

        #endregion

        public void НажатиеНаКнопкуПереслать()
        {
            //выбираем адресатов
            List<ReferenceObject> users = ВыбратьПользователей();
            //если выбрали 
            if (users != null && users.Count > 0)
            {
                //заносим их в доп доступ
                RegistryControlCard registryControlCard = new RegistryControlCard(Context.ReferenceObject);
                foreach (var user in users)
                {
                    registryControlCard.ДобавитьВДополнительныйДоступ(user);
                }
                if (registryControlCard.ПрименитьИзменения())
                {
                    //генерируем письмо
                    HTML_MailMessageI message = new HTML_MailMessageI();
                    string text = "Вам предоставлен доступ к объекту справочника Канцелярия";
                    message.Subject = text;
                    message.BodyText = text;
                    message.AddAttachment(registryControlCard.ReferenceObject);
                    List<IMailRecipient> mailRecipients = (from user in users
                                                           select new MailRecipient(user as UserReferenceObject)).Select(mr => (IMailRecipient)mr).ToList();
                    message.MailRecipients = mailRecipients;

                    //показываем письмо
                    message.Show();
                }
                else MessageBox.Show("Не удалось добавить адресатов в дополнительный доступ.");
            }
            //если не выбрали то конец
        }

        /// <summary>
        /// возвращает список выбранных пользователей
        /// </summary>
        /// <returns></returns>
        private List<ReferenceObject> ВыбратьПользователей()
        {
            List<ReferenceObject> result = null;
            UIMacroContext uiContext = Context as UIMacroContext;
            if (uiContext != null)
            {
                using (var dialog = TFlex.DOCs.Model.Plugins.ObjectCreator.CreateObject<ISelectROFromReferencesDialog>())
                {
                    dialog.Text = "Выберите адресатов";
                    dialog.Initialize(true, new Reference[] { Context.Connection.References.Users });
                    // Первый аргумент позволяет задать множественный выбор (появляется правая часть со списком выбранных объектов).
                    // Второй аргумент отвечает за те справочники, из которых выбираются объекты. Если нужны все справочники - передаем null.
                    // Если нужно из нескольких справочников выбрать один, отображаемый при открытии окна, то задаем свойство dialog.SelectedReference

                    if (dialog.ShowDialog(uiContext.OwnerWindow) == DialogOpenResult.Ok)
                    {

                        result = dialog.SelectedReferenceObjectList.ToList(); // Получаем выбранные объекты
                    }
                }
            }
            return result;
        }

        struct RegistryControlCardType
        {
            public const string Internal = "Внутренняя";
            public const string Outgoing = "Исходящая";
            public const string Incoming = "Входящая";
        }

        public void НажатиеНаКнопкуОтправитьПолучателям()
        {
            ReferenceObject current_ro = (ReferenceObject)ТекущийОбъект;//Context.ReferenceObject; //текущий объект
            switch (current_ro.Class.Name)
            {
                case RegistryControlCardType.Internal:
                    SendInternalRegistryControlCard(current_ro);
                    break;
                case RegistryControlCardType.Outgoing:
                    SendOutgoingRegistryControlCard(current_ro);
                    break;
                case RegistryControlCardType.Incoming:
                    SendIncomingRegistryControlCard(current_ro);
                    break;
                default:
                    MessageBox.Show("Нет подходящего метода для типа " + current_ro.Class.Name, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }

        }

        public void SendIncomingRegistryControlCard(ReferenceObject current_ro)
        {

            ServerConnection servConnect = (Context as MacroContext).Connection;
            ТекущийОбъект.Параметр["Статус прохождения"] = "1";

            RegistryControlCard registryControlCard = new RegistryControlCard(current_ro);
            //UserReference polzovateli = new UserReference(servConnect);

            if (registryControlCard.Responsible == null)
            {
                MessageBox.Show("Не указан ответственный!", "Ошибка!",
                       MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            ДиалогВвода диалог = СоздатьДиалогВвода("Задайте параметры задания");
            //диалог.ДобавитьДату("Дата начала", DateTime.Now);
            диалог.ДобавитьДату("Срок выполнения", DateTime.Now.AddDays(3));
            диалог.ДобавитьСтроковое("Текст задания", "", true);
            диалог.ДобавитьВыборИзСправочника("Отправка от имени", "Группы и пользователи");


            DateTime task_begin = DateTime.Now;
            DateTime task_end = DateTime.Now;
            string task_text = "";
            User task_fromwho = null;
            ReferenceObject rot_task_fromwho = null;

            if (диалог.Показать())
            {
                task_end = диалог["Срок выполнения"];
                task_text = диалог["Текст задания"];

                rot_task_fromwho = (ReferenceObject)диалог["Отправка от имени"];

                task_fromwho = rot_task_fromwho as User;

                if (task_fromwho == null)
                    MessageBox.Show("Указание группы пользователей в поле <<Отправка от имени>> запрещено!\nУкажите конкретного пользователя", "Ошибка!",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            else
            {
                Сообщение("Действия диалога ввода", "Отмена!");
                return;
            }

            User otvetstvenniy = registryControlCard.Responsible as User;

            SendIncomingMailForInform(registryControlCard, task_end, task_text);

            MailTask task = new MailTask(servConnect); //Создаем новое задание ответственному
            string taskSubj = "Вы назначены ответственным на входящее N " + registryControlCard.RegistryNumber + " от " + registryControlCard.From;
            task.Subject = taskSubj;
            string taskMsg = "Вам направлена входящая корреспонденция, на которую Вы назначены Ответственным за исполнение."
                + "<br>Вам необходимо выполнить следующие действия:<br>"
                + "<br>1. Для подтверждения принятия задачи нажмите кнопку \"Принять\" в верхнем левом углу  окна задания."
                + "<br>2. Выполните задание, описанное в содержании."
                + "<br>3. Перейдите к вложенному объекту (иконка справа от задания), "
                + "во вкладку \"Входящие - Все пользователи\" и нажмите кнопку \"Подтверждаю исполнение\".<br>"
                + "<br>4. Нажмите кнопку \"Завершить\" в верхнем левом углу окна задания.<br>";
            task.SetBody(CreateHTMLMailBodyText(taskSubj, taskMsg, registryControlCard), MailBodyType.Html);


            task.Executors.Add(new MailTaskExecutor(otvetstvenniy)); //Добавляем исполнителя задания
            foreach (var item in registryControlCard.Documents) //добавляем вложения
            {
                task.Attachments.Add(new ObjectAttachment(item));
            }
            User controller = registryControlCard.Checker as User;
            task.Controller = new MailUser(controller);//Добавляем контроллера задания
                                                       // MessageBox.Show();
                                                       //Письмо ответственному на внешнюю почту
            if (otvetstvenniy.MailSendType > 0 && otvetstvenniy.Email != null)
            {

                //MailMessage messageTask = new MailMessage(servConnect.Mail.DOCsAccount); //Создаем новое сообщение
                string textSubj = "Вы назначены ответственным на входящее N " + registryControlCard.RegistryNumber + " от " + registryControlCard.From;
                string textMsg = "Вам направлена входящая корреспонденция, на которую Вы назначены Ответственным за исполнение." +
                    "<br>Откройте АСУ \"Динамика\" и выполните следующие действия:<br>"
                    + "<br>1. Для подтверждения принятия задачи нажмите кнопку \"Принять\" в верхнем левом углу  окна задания."
                    + "<br>2. Выполните задание, описанное в содержании."
                    + "<br>3. Перейдите к вложенному объекту (иконка справа от задания), "
                    + "    во вкладку \"Входящие - Все пользователи\" и нажмите кнопку \"Подтверждаю исполнение\"."
                    + "<br>4. Нажмите кнопку \"Завершить\" в верхнем левом углу окна задания.<br>"
                    + "<br>Задание: " + task_text + "<br>Срок выполнения:  " + task_end.ToString("dd.MM.yyyy") + "<br>";
                List<ReferenceObject> mailAttachments = new List<ReferenceObject> { registryControlCard.ReferenceObject };
                mailAttachments.AddRange(registryControlCard.Documents);
                SendHTMLMessage(textSubj, textMsg, new List<ReferenceObject> { otvetstvenniy }, mailAttachments);
            }
            task.StartDate = DateTime.Now; //Дата начала выполнения задания
            task.EndDate = task_end; //Срок выполнения задания
                                     //task.CheckDate = DateTime.Now.AddDays(5); //Контрольный срок задания

            task.Priority = MailTaskPriority.Hight; //Важность задания

            //Прикрепляем к заданию объект
            task.Attachments.Add(new ObjectAttachment(registryControlCard.ReferenceObject));
            if (task_fromwho != null) task.OnBehalf = new MailUser(task_fromwho as User); //Добавляем поле "от имени"
            task.Send(); //Отправляем задание

            ТекущийОбъект.Параметр["Примечание"] = task_text;
            ChangeStage(registryControlCard, StageName.ApprovedByDirector);

            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true); // Сохраняем изменения и закрываем диалог
        }

        bool ChangeStage(RegistryControlCard rcc, string stageName)
        {
            Stage stage = TFlex.DOCs.Model.Stages.Stage.Find(Connection, stageName);
            return ((stage.Change(new List<ReferenceObject> { rcc.ReferenceObject })).Count > 0);
        }

        private void SendIncomingMailForInform(RegistryControlCard registryControlCard, DateTime task_end, string task_text)
        {
            User otvetstvenniy = registryControlCard.Responsible as User;
            string textSubj = "Вам перенаправлена входящая корреспонденция N " + registryControlCard.RegistryNumber + " от " + registryControlCard.From;
            string textMsg = "<br>1. Одинарным кликом откройте вложенный объект, который находится справа от иконки \"Сообщение\""
+ "<br>2. В открывшемся диалоге перейдите на вкладку \"Журнал ознакомления\""
+ "<br>3. Нажмите на кнопку \"Оставить запись об ознакомлении\" <br>4. В открывшемся диалоге подтвердите ознакомление нажав кнопку ОК<br>"
+ "<br>Поставленное задание: " + task_text
+ "<br>Срок выполнения:  " + task_end.ToString("dd.MM.yyyy")
+ "<br>Ответственный:   " + otvetstvenniy.ToString();

            // получение объектов списка или объектов по связи 1:N, N:N, на любой справочник
            List<ReferenceObject> mailRecipients = registryControlCard.SignedTo.ToList();

            List<ReferenceObject> mailAttachments = new List<ReferenceObject> { registryControlCard.ReferenceObject };
            mailAttachments.AddRange(registryControlCard.Documents);

            SendHTMLMessage(textSubj, textMsg, mailRecipients, mailAttachments);
        }

        /// <summary>
        /// Осуществляет рассылку внутренним пользователям
        /// </summary>
        /// <param name="current_ro"></param>
        private void SendInternalRegistryControlCard(ReferenceObject current_ro)
        {
            RegistryControlCard registryControlCard = new RegistryControlCard(current_ro);


            if (registryControlCard.Performer == null)
            {
                MessageBox.Show("Не указан исполнитель!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //создаём список адресатов
            List<MailRecipient> mailRecipients = registryControlCard.GetUsersForAccess().Select(u => new MailRecipient(u as UserReferenceObject)).ToList();
            //добавляется в GetUsersForAccess()
            //mailRecipients.Add(new ШаблонДляМакроса.MailRecipient( registryControlCard.Performer as UserReferenceObject));
            mailRecipients = mailRecipients.Distinct().ToList();
            List<MailRecipient> mailCopyRecipients = new List<Cancelaria.MailRecipient>();
            if (!ВыбратьАдресатов(ref mailRecipients, ref mailCopyRecipients)) return;
            //Сохранить объект
            registryControlCard.ReferenceObject.ApplyChanges();
            StringBuilder text = new StringBuilder();
            if (mailRecipients.Count > 0)
            {
                //Отправить письмо пользователям из дополнительного доступа и Исполнителю
                ОтправитьПисьмоОНовойЗаписиВКанцелярии(registryControlCard, mailRecipients);
                text.AppendLine("Осуществлена отправка следующим получателям:");
                foreach (var user in mailRecipients)
                    text.AppendLine(user.Name);
            }

            ChangeStage(registryControlCard, StageName.ApprovedByDirector);

            text.AppendLine("Стадия изменена на \"Подтверждено директором\".");
            //Вывести сообщение об отправке
            MessageBox.Show(text.ToString());
        }

        /// <summary>
        /// Осуществляет рассылку исходящей корреспонденции
        /// </summary>
        /// <param name="current_ro"></param>
        private void SendOutgoingRegistryControlCard(ReferenceObject current_ro)
        {
            RegistryControlCard registryControlCard = new RegistryControlCard(current_ro);

            //проверяем внешних адресатов на наличие Email
            IEnumerable<Organization> toOrganizations = registryControlCard.Get_ToOrganizations();
            //if (!IsOrganizationsEmailsAreCorrect(toOrganizations)) return;

            if (registryControlCard.Performer == null)
            {
                MessageBox.Show("Не указан исполнитель!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //получаем список контактов в организациях
            List<MailRecipient> mailRecipients = GetMailRecipientsFromOrganizations(toOrganizations);
            List<MailRecipient> mailCopyRecipients = new List<MailRecipient>();// { new MailRecipient(registryControlCard.Performer as UserReferenceObject) };
            List<MailRecipient> extraAccess = registryControlCard.GetUsersForAccess().Select(r => new MailRecipient(r as UserReferenceObject)).ToList();
            mailCopyRecipients.AddRange(extraAccess);
            //показываем диалог выбора адресатов
            if (!ВыбратьАдресатов(ref mailRecipients, ref mailCopyRecipients)) return;

            HTML_MailMessageI mailMessage = new HTML_MailMessageI();
            //добавляем адресатов
            mailMessage.MailCopyRecipients = mailCopyRecipients.Select(mr => (IMailRecipient)mr).ToList();
            mailMessage.MailRecipients = mailRecipients.Select(mr => (IMailRecipient)mr).ToList();
            //создаём список вложений
            List<ReferenceObject> mailAttachments = registryControlCard.Documents.ToList();
            mailAttachments.Add(registryControlCard.ReferenceObject);
            foreach (var item in mailAttachments)
            {
                mailMessage.AddAttachment(item);
            }

            //создаём тему
            string subject = string.Format(@"Письмо от АО ЦНТУ Динамика №{0}, от {1} {2}", registryControlCard.RegistryNumber, ((DateTime)(registryControlCard.RegistryDate)).ToShortDateString(), registryControlCard.Content);
            if (!string.IsNullOrWhiteSpace(subject)) subject += "; ";
            if (!string.IsNullOrWhiteSpace(registryControlCard.RegistryNumber)) subject += registryControlCard.RegistryNumber + "; ";
            if (registryControlCard.RegistryDate != null) subject += "от " + ((DateTime)(registryControlCard.RegistryDate)).ToShortDateString() + "; ";
            mailMessage.Subject = subject;

            //создаём текст письма
            mailMessage.BodyText = string.Format(@"<p>Здравствуйте!</p><p>Во вложении {0}</p> 
<p>С уважением,
</p>
<p>
АО ЦНТУ «Динамика»<br>
Тел.: +7 (495) 2760009<br>
<a href = www.dinamika-avia.ru>www.dinamika-avia.ru</a></p>", string.Format(@"письмо от АО ЦНТУ ""Динамика"" №{0}, от {1} {2}", registryControlCard.RegistryNumber, ((DateTime)(registryControlCard.RegistryDate)).ToShortDateString(), registryControlCard.Content));


            ///////////////////////

            StringBuilder text = new StringBuilder();
            if (mailRecipients.Count > 0 || mailCopyRecipients.Count > 0)
            {
                //отправляем письмо
                mailMessage.Send();
                text.AppendLine("Осуществлена отправка следующим получателям:");
                foreach (var user in mailRecipients)
                    text.AppendLine(user.Name);
                foreach (var user in mailCopyRecipients)
                    text.AppendLine(user.Name.ToString());
            }

            ReferenceObject mailShipmentType = GetMailShipmentByName("Электронная почта");
            registryControlCard.ReferenceObject.BeginChanges();
            registryControlCard.ReferenceObject.SetLinkedObject(new Guid("2bb8589f-89a6-4666-8dae-a71d3180af10"), mailShipmentType);
            registryControlCard.ReferenceObject[new Guid("53b077d7-4d0a-4a2e-9541-28c790f81c1b")].Value = DateTime.Now;
            registryControlCard.ReferenceObject[new Guid("7d119656-b7d1-4edc-a110-b4984bb9cf8d")].Value = "Отправлено по электронной почте";
            registryControlCard.ReferenceObject.EndChanges();

            ChangeStage(registryControlCard, StageName.ApprovedByDirector);

            text.AppendLine("Стадия изменена на \"Подтверждено директором\".");
            ЗакрытьДиалогСвойств();
            //Вывести сообщение об отправке
            MessageBox.Show(text.ToString());
        }

        string CreateHTMLMailBodyText(string textSubj, string textMsg, RegistryControlCard registryControlCard)
        {
            StringBuilder strBld = new StringBuilder();
            ServerConnection servConnect = (Context as MacroContext).Connection;
            string recipients = "";
            List<ReferenceObject> resps = registryControlCard.SignedTo.ToList();
            if (resps != null && resps.Count > 0)
                foreach (var item in resps)
                {
                    if (recipients == "") recipients = item.ToString();
                    else { recipients += "<br>" + item.ToString(); }
                }

            String responsible = registryControlCard.Responsible.ToString();

            strBld.Append(String.Format(@"<html>
	<head>
		<meta charset=""utf-8"">
		<title>{0}</title>
	</head>
    <body><span style='font-size:12.0pt; font-family:""Arial"",""sans-serif"";mso-fareast-font-family:""Times New Roman"";
mso-fareast-theme-font:minor-fareast;mso-fareast-language:RU;mso-no-proof:yes'
    <p>Здравствуйте!</p>
    <p>Это автоматически созданное письмо.<o:p></o:p></p>
    <p> {1} </p>

			<table  cellpadding=""0"" cellspacing=""0"" border=""1"" width=""99%"">
			
			 <tr>
			    <th nowrap bgcolor=#e1e1e1> Рег. № </th>
			    <th nowrap bgcolor=#e1e1e1> Дата </th>
			    <th nowrap bgcolor=#e1e1e1> Откуда </th>
				<th nowrap bgcolor=#e1e1e1> Содержание </th>
				<th nowrap bgcolor=#e1e1e1> Ответственный </th>
				<th nowrap bgcolor=#e1e1e1> Получатели </th>
			 </tr>", textSubj.ToLower(), textMsg));
            string link = HTML_MailMessageI.GetLinkFor(registryControlCard.ReferenceObject);

            strBld.Append(String.Format(@"
<tr>
<td style=align=""center"">{0}</td>
<td>{1}</td>
<td>{2}</td>
<td style=align=""center""><a href={3}>{4}</a></td>
<td>{5}</td>
<td>{6}</td>
</tr>",
                                        registryControlCard.RegistryNumber,
                                        ((DateTime)(registryControlCard.RegistryDate)).ToString("dd.MM.yyyy"),
                                        registryControlCard.From,
                                        link,
                                        registryControlCard.Content,
                                        responsible,
                                        recipients
                                       ));
            //}
            strBld.Append(String.Format(@"</table>
        </tr>
    </table>
    <!-- End Container Table -->
<p class=MsoNormal>
<span style='font-size:12.0pt;font-family:""Arial"",""sans-serif"";mso-fareast-font-family:
""Times New Roman"";mso-fareast-theme-font:minor-fareast;mso-fareast-language:
RU;mso-no-proof:yes'>Если вы получили это письмо ошибочно или у вас возникли вопросы обращайтесь в
<a href=""mailto:Служба%20внедрения%20и%20сопровождения%20систем%20автоматизации"">Отдел автоматизации</a>
 по тел. 1180, 1200, 1204.<o:p></o:p></p>
</span></p>"));
            string textMsgInner = strBld.ToString() +
                @"<p>
                    <img src=""http://portal.dinamika-avia.ru/include/logo.2186.png""  alt=""АСУ ""Динамика"""">
                  </p>
			</span></body>
			</html>";
            return textMsgInner;
        }

        private ReferenceObject GetMailShipmentByName(string mailShipmentName)
        {
            Guid mailShipmentReferenceGuid = new Guid("2e6a9419-e008-4f45-901a-30ce2caff481");
            ReferenceInfo mailShipmentReferenceInfo = References.Connection.ReferenceCatalog.Find(mailShipmentReferenceGuid);
            // создание объекта для работы с данными
            Reference reference = mailShipmentReferenceInfo.CreateReference();
            ReferenceObject mailShipment = reference.Find(reference.ParameterGroup.Parameters.Find(new Guid("e2bb86d2-c1aa-4eba-8810-a5e0b94f6493")), mailShipmentName).FirstOrDefault();
            return mailShipment;
        }

        public void ЗакрытьДиалогСвойств()
        {
            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true);
        }

        private static List<MailRecipient> GetMailRecipientsFromOrganizations(IEnumerable<Organization> toOrganizations)
        {
            List<MailRecipient> mailRecipients = new List<Cancelaria.MailRecipient>();
            foreach (var organization in toOrganizations)
            {
                mailRecipients.Add(new Cancelaria.MailRecipient(organization.Name, organization.Email));
                foreach (var contact in organization.Contacts)
                {
                    mailRecipients.Add(new MailRecipient(contact.FullName + " (" + contact.Email + ")", contact.Email));
                }
            }
            return mailRecipients;
        }

        private static bool IsOrganizationsEmailsAreCorrect(IEnumerable<Organization> toOrganizations)
        {
            StringBuilder errorsMessages = new StringBuilder();
            if (toOrganizations.Count() == 0)
            {
                errorsMessages.AppendLine("Не задан адресат! Заполните поле Куда адресован.");
            }

            foreach (var organisation in toOrganizations)
            {
                if (!HTML_MailMessageI.IsValidEmail(organisation.Email))
                {
                    errorsMessages.AppendLine("У организации " + organisation.Name + ", указан некорректный E-mail адрес");
                }
            }
            if (errorsMessages.Length > 0)
            {
                MessageBox.Show(errorsMessages.ToString(), "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        //Если диалог закрыт кнопкой Ок то вернуть true, в противном случае false
        /*private bool ВыбратьАдресатов(ref List<ReferenceObject> mailRecipients)
        {
            bool dialogResult = false;
            ДиалогВвода диалог = СоздатьДиалогВвода("Выберите адресатов");
            foreach (var mailRecipient in mailRecipients)
                диалог.ДобавитьФлаг(mailRecipient.ToString(), true);
            if (диалог.Показать())
            {
                List<ReferenceObject> result = new List<ReferenceObject>();
                foreach (var mailRecipient in mailRecipients)
                { if (диалог[mailRecipient.ToString()]) result.Add(mailRecipient); }
                mailRecipients = result;
                //if(mailRecipients.Count > 0)
                dialogResult = true;
            }

            return dialogResult;
        }*/

        private bool ВыбратьАдресатов(ref List<MailRecipient> mailRecipients, ref List<MailRecipient> mailCopyRecipients)
        {
            bool dialogResult = false;
            ДиалогВвода диалог = СоздатьДиалогВвода("Выберите адресатов");
            диалог.ДобавитьГруппу("Кому:");
            foreach (var mailRecipient in mailRecipients)
            {
                bool isEmailCorrect = HTML_MailMessageI.IsValidEmail(mailRecipient.Email);
                диалог.ДобавитьФлаг(mailRecipient.Name, isEmailCorrect);
                диалог.УстановитьДоступностьЭлемента(mailRecipient.Name, isEmailCorrect);
                if (!isEmailCorrect)
                    диалог.ДобавитьКомментарий(mailRecipient.Name, "E-Mail адрес не корректен!");
            }
            if (mailCopyRecipients.Count > 0)
            {
                диалог.ДобавитьГруппу("Копия:");
                foreach (var mailRecipient in mailCopyRecipients)
                {
                    bool isEmailCorrect = HTML_MailMessageI.IsValidEmail(mailRecipient.Email);
                    диалог.ДобавитьФлаг(mailRecipient.Name, isEmailCorrect);
                    диалог.УстановитьДоступностьЭлемента(mailRecipient.Name, isEmailCorrect);
                    if (!isEmailCorrect)
                        диалог.ДобавитьКомментарий(mailRecipient.Name, "E-Mail адрес не корректен!");
                }
            }
            диалог.УстановитьРазмер(800, 0);

            if (диалог.Показать())
            {
                List<MailRecipient> result = new List<MailRecipient>();
                foreach (var mailRecipient in mailRecipients)
                {
                    //MessageBox.Show(диалог[mailRecipient.Name].ToString() + "\n" + mailRecipient.Name);
                    if (диалог[mailRecipient.Name]) result.Add(mailRecipient);
                }
                mailRecipients = result;

                List<MailRecipient> resultCopy = new List<MailRecipient>();
                foreach (var mailRecipient in mailCopyRecipients)
                {
                    //MessageBox.Show(диалог[mailRecipient.Name].ToString() +"\n" + mailRecipient.Name);
                    if (диалог[mailRecipient.Name]) result.Add(mailRecipient);
                }
                mailCopyRecipients = resultCopy;
                //if(mailRecipients.Count > 0)
                dialogResult = true;

            }

            return dialogResult;
        }

        private void ОтправитьПисьмоОНовойЗаписиВКанцелярии(RegistryControlCard registryControlCard, List<MailRecipient> mailRecipients)
        {
            HTML_MailMessageI mailMessage = new HTML_MailMessageI();
            mailMessage.MailRecipients = mailRecipients.Select(mr => (IMailRecipient)mr).ToList();
            //создаём список вложений
            List<ReferenceObject> mailAttachments = registryControlCard.Documents.ToList();
            mailAttachments.Add(registryControlCard.ReferenceObject);
            foreach (var item in mailAttachments)
            {
                mailMessage.AddAttachment(item);
            }

            //создаём тему
            string subject = "";

            if (registryControlCard.RegistryJournal != null && registryControlCard.RegistryJournal.Name.Contains("Журнал регистрации служебных записок"))
            {
                subject += "Служебная записка";
            }
            else {
                subject += registryControlCard.StorageFolder;
            }

            if (!string.IsNullOrWhiteSpace(subject)) subject += "; ";
            if (!string.IsNullOrWhiteSpace(registryControlCard.RegistryNumber)) subject += registryControlCard.RegistryNumber + "; ";
            if (registryControlCard.RegistryDate != null) subject += "от " + ((DateTime)(registryControlCard.RegistryDate)).ToShortDateString() + "; ";
            subject += registryControlCard.Content;
            mailMessage.Subject = subject;
            //создаём текст письма
            mailMessage.BodyText = string.Format(@"<p>Здравствуйте!</p><p>Вам направлена новая запись канцелярии - {0}, см. вложение. </p>", subject);
            mailMessage.Send();
        }

        #region Отправить карточку канцелярии о служебной записке На Контроль в отдел планирования и контроля при нажатии на кнопку
        public void ОтправитьНаКонтроль()
        {
            ReferenceObject current_ro = (ReferenceObject)ТекущийОбъект;//Context.ReferenceObject; //текущий объект
            RegistryControlCard служебнаяЗаписка = new RegistryControlCard(current_ro);
            //Добавить в дополнительный доступ НачальникОтделаПланированияИКонтроля и НачальникГруппыКонтроляИАналитики
            служебнаяЗаписка.ДобавитьВДополнительныйДоступ(НачальникОтделаПланированияИКонтроля);
            служебнаяЗаписка.ДобавитьВДополнительныйДоступ(НачальникГруппыКонтроляИАналитики);
            //Сохранить объект
            служебнаяЗаписка.ПрименитьИзменения();
            //Отправить письмо НачальникОтделаПланированияИКонтроля и НачальникГруппыКонтроляИАналитики о новой служебке
            ОтправитьПисьмоОНовойСлужебнойЗаписке(служебнаяЗаписка, new List<ReferenceObject>() { НачальникОтделаПланированияИКонтроля, НачальникГруппыКонтроляИАналитики });
            //Вывести сообщение об отправке "Отдел контроля - проинформирован"
            MessageBox.Show("Отдел планирования и контроля оповещён.");
        }

        private void ОтправитьПисьмоОНовойСлужебнойЗаписке(RegistryControlCard служебнаяЗаписка, List<ReferenceObject> списокАдресатов_ro)
        {
            string заголовокПисьма = "Новая служебная записка.";
            string текстПисьма = СформироватьТекстПисьма(служебнаяЗаписка);
            List<ReferenceObject> вложения = new List<ReferenceObject>() { служебнаяЗаписка.ReferenceObject };
            SendHTMLMessage(заголовокПисьма, текстПисьма, списокАдресатов_ro, вложения);
        }

        private string СформироватьТекстПисьма(RegistryControlCard служебнаяЗаписка)
        {
            StringBuilder текстПисьма = new StringBuilder();
            текстПисьма.Append("В канцелярии появилась новая служебная записка.");
            string link = String.Format(@"docs://{0}/OpenReferenceWindow/?refId={1}&objId={2}",
                                        Connection.ConnectionParameters.GetServerAddress(), служебнаяЗаписка.ReferenceObject.Reference.Id, служебнаяЗаписка.ReferenceObject.SystemFields.Id);
            //MessageBox.Show("Ссылка - " + String.Format(@"<a href ={0}>{1}</a>", link, служебнаяЗаписка.ОбъектСправочника_ro.ToString()) + ".");
            текстПисьма.AppendLine("<p>Ссылка - " + String.Format(@"<a href ={0}>{1}</a>", link, служебнаяЗаписка.ReferenceObject.ToString()) + @".</p>");
            return текстПисьма.ToString();
        }
        #endregion
        public void SendHTMLMessage(string textSubj, string textMsg, List<ReferenceObject> mailRecipients, List<ReferenceObject> attachments = null)
        {
            ServerConnection servConnect = (Context as MacroContext).Connection;
            //Создаем новое сообщение для внутренней почты
            MailMessage message = new MailMessage(servConnect.Mail.DOCsAccount);

            //прикрепляем вложения если они есть
            if (attachments != null)
            {
                foreach (ReferenceObject item in attachments)
                {
                    message.Attachments.Add(new ObjectAttachment(item));
                }
            }
            StringBuilder strBld = new StringBuilder();

            strBld.Append(String.Format(@"<html>
	<head>
		<meta charset=""utf-8"">
		<title>{0}</title>
	</head>
    <body><span style='font-size:12.0pt; font-family:""Arial"",""sans-serif"";mso-fareast-font-family:""Times New Roman"";
mso-fareast-theme-font:minor-fareast;mso-fareast-language:RU;mso-no-proof:yes'
    <p>Здравствуйте!</p>

    <p> {1} </p>
	    ", textSubj.ToLower(), textMsg));

            strBld.Append(String.Format(@"</table>
        </tr>
    </table>
    <!-- End Container Table -->
<p class=MsoNormal>
<span style='font-size:12.0pt;font-family:""Arial"",""sans-serif"";mso-fareast-font-family:
""Times New Roman"";mso-fareast-theme-font:minor-fareast;mso-fareast-language:
RU;mso-no-proof:yes'>Если вы получили это письмо ошибочно или у вас возникли вопросы обращайтесь в
<a href=""mailto:Служба%20внедрения%20и%20сопровождения%20систем%20автоматизации"">Отдел автоматизации</a>
 по тел. 1180, 1200, 1204<o:p></o:p></p>
<p>или по E-mail: <a href=""mailto:lipaev@dinamika-avia.ru"">lipaev@dinamika-avia.ru</a>
<o:p></o:p></span></p>"));
            string textMsgInner = strBld.ToString() + String.Format(@"<p><img src=""data:image/png;base64,{0}  alt=""АСУ ""Динамика""""></p>
			</span></body>
			</html>", Logo_Base64);
            string textMsgOuter = strBld.ToString() + "<//span><//body><//html>";
            //<p><img src=""data:image/png;base64,{0}  alt=""АСУ ""Динамика""""></p>

            message.SetBody(textMsgInner, MailBodyType.Html);

            message.Subject = CleanString(textSubj);
            /*message.To.Add(new EMailAddress("lipaev@dinamika-avia.ru"));
                    message.To.Add(new MailUser(mailuser));
             */
            //Добавление адресатов
            List<ReferenceObject> tempUsers = new List<ReferenceObject>();
            foreach (UserReferenceObject mailgroup in mailRecipients)
            {
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

            foreach (UserReferenceObject user in users.Distinct())
            {
                User mailuser = user as User;
                message.To.Add(new MailUser(mailuser));

                if (mailuser.Email != null && !mailuser.Email.IsEmpty)
                {
                    message.To.Add(new EMailAddress(mailuser.Email.ToString()));
                }
            }
            message.Send(); //Отправляем сообщение
        }
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
        #region Кнопка для канцелярии Создать запись в Библиотеке ОРД

        public void СоздатьЗаписьВБилиотекеОРД()
        {
            ReferenceObject current = Context.ReferenceObject;
            current.ApplyChanges();
            RegistryControlCard rc = new RegistryControlCard(current);
            if (rc.HasBibleORDItem())
            {
                MessageBox.Show("Для текущего объекта уже создана запись в Библиотеке ОРД!");
                return;
            }
            BiblORDRecord bord = new BiblORDRecord();

            bord.SetAccessLevel(rc.GetAccessLevel());
            bord.AddToAccessList(rc.GetUsersForAccess());
            bord.SetName(rc.Content);
            bord.SetFileDocument(GetTheFileWithLatestCreationDate(rc.GetAllLinkedFiles()));
            ReferenceObject newRO = bord.GetNewBORDReferenceObject();
            bord.LinkWithRCC(current);
            RefObj ro = RefObj.CreateInstance(newRO, Context);
            if ((Context as MacroContext).ShowObjectPropertyDialog(ro, "")) Закрыть();
        }

        private ReferenceObject GetTheFileWithLatestCreationDate(IEnumerable<ReferenceObject> files)
        {
            ReferenceObject result = null;
            DateTime latestDate = DateTime.MinValue;
            foreach (var file in files)
            {
                if (file.SystemFields.CreationDate.CompareTo(latestDate) < 0) continue;
                result = file;
                latestDate = file.SystemFields.CreationDate;
            }
            return result;
        }



        #endregion


        public void Закрыть()
        {
            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true);
        }

        #region Группы и пользователи
        private static Guid Users_Руководитель_UsersGroup_link_GUID = new Guid("fdb41549-2adb-40b0-8bb5-d8be4a6a8cd2"); //Guid связи "Руководитель" типа "Группа пользователей" справочника - "Группы и пользователи"
        private static Guid Users_ОтделПланированияИКонтроля_item_GUID = new Guid("0fd02ae4-68aa-480b-80f8-2904b2720451"); //Guid объекта "Отдел планирования и контроля" справочника - "Группы и пользователи"
        private static Guid Users_ГруппаКонтроляИАналитики_item_GUID = new Guid("a44b32b0-b9ec-439b-b42d-0678da7dc882"); //Guid объекта "Группа контроля и аналитики" справочника - "Группы и пользователи"
        #endregion

        //поля
        private Reference _userReference;
        private Reference _RegistryControlCardReference;
        private ReferenceInfo _userReferenceInfo;
        private ReferenceInfo _RCCReferenceInfo;
        private ParameterInfo _PD_PersonCodeFrom1C;
        private ReferenceObject _НачальникОтделаПланированияИКонтроля;
        private ReferenceObject _НачальникГруппыКонтроляИАналитики;

        public const int LogoFileID = 68387;
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

        public static ServerConnection Connection
        {
            get { return ServerGateway.Connection; }
        }

        #region Справочники

        private Reference GetReference(ref Reference reference, ReferenceInfo referenceInfo)
        {
            if (reference == null)
                reference = referenceInfo.CreateReference();


            return reference;
        }

        /// <summary>
        /// Справочник "Группы и пользователи"
        /// </summary>
        private Reference UserReference
        {
            get
            {
                if (_userReference == null)

                    return GetReference(ref _userReference, UserReferenceInfo);

                return _userReference;
            }
        }

        /// <summary>
        /// Справочник "Регистрационно-контрольные карточки"
        /// </summary>
        private Reference RCCReference
        {
            get
            {
                if (_RegistryControlCardReference == null)

                    return GetReference(ref _RegistryControlCardReference, RCCReferenceInfo);

                return _RegistryControlCardReference;
            }
        }

        #endregion
        private ReferenceInfo UserReferenceInfo
        {
            get { return GetReferenceInfo(ref _userReferenceInfo, reference_Users_Guid); }
        }

        private ReferenceInfo RCCReferenceInfo
        {
            get { return GetReferenceInfo(ref _RCCReferenceInfo, reference_RCC_Guid); }
        }

        private static ReferenceInfo GetReferenceInfo(ref ReferenceInfo referenceInfo, Guid referenceGuid)
        {
            if (referenceInfo == null)
                referenceInfo = Connection.ReferenceCatalog.Find(referenceGuid);

            return referenceInfo;
        }

        private ParameterInfo GetParameterInfo(ref ParameterInfo parameterInfo, ReferenceInfo referenceInfo, Guid parameterGuid)
        {
            if (parameterInfo == null)
                parameterInfo = referenceInfo.Description.OneToOneParameters.Find(parameterGuid);

            return parameterInfo;
        }

        private ReferenceObject НачальникОтделаПланированияИКонтроля
        {
            get
            {
                if (_НачальникОтделаПланированияИКонтроля == null)
                    _НачальникОтделаПланированияИКонтроля = (UserReference.Find(Users_ОтделПланированияИКонтроля_item_GUID) as UsersGroup).GetObject(Users_Руководитель_UsersGroup_link_GUID);
                return _НачальникОтделаПланированияИКонтроля;
            }
        }
        private ReferenceObject НачальникГруппыКонтроляИАналитики
        {
            get
            {
                if (_НачальникГруппыКонтроляИАналитики == null)
                    _НачальникГруппыКонтроляИАналитики = (UserReference.Find(Users_ГруппаКонтроляИАналитики_item_GUID) as UsersGroup).GetObject(Users_Руководитель_UsersGroup_link_GUID);
                return _НачальникГруппыКонтроляИАналитики;
            }
        }
    }

    struct StageName
    {
        public const string ApprovedByDirector = "Канцелярия. Подтверждено Директором";
        public const string WaitForApprove = "Канцелярия. Ждет подтверждения";
    }

    /// <summary>
    /// Для работы с регистрационно-контрольными карточками Канцелярии
    /// ver 1.2
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

        RegistryJournal _RegistryJournal;
        public RegistryJournal RegistryJournal
        {
            get
            {
                if (_RegistryJournal == null)
                {
                    ReferenceObject ro = ReferenceObject.GetObject(RCC_link_RegistryJournal_GUID);
                    if (ro != null) _RegistryJournal = new RegistryJournal(ro);
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

    public class Organization : IMailRecipient
    {
        public Organization(ReferenceObject referenceObject)
        {
            this.ReferenceObject = referenceObject;
            isNewOrganization = referenceObject.Reference.ParameterGroup.ReferenceInfo.Guid.ToString() == "cb7442b8-5b69-490a-87da-937ac39f0a32";
        }

        private bool isNewOrganization;

        internal ReferenceObject ReferenceObject;

        List<Contact> _Contacts;
        public IEnumerable<Contact> Contacts
        {
            get
            {
                if (_Contacts == null)
                {
                    _Contacts = new List<Contact>();
                    Guid contactsGuid;
                    if (isNewOrganization)
                    {
                        contactsGuid = new Guid("46dc7f9d-d17e-4feb-8b17-951728325e1d");
                    }
                    else
                    {
                        contactsGuid = new Guid("9454188c-0a39-4fad-aeee-3e6dcd1c89c4");
                    }
                    foreach (var item in this.ReferenceObject.GetObjects(contactsGuid))
                    {
                        _Contacts.Add(FactoryContact.CreateContact(item));
                    }
                }
                return _Contacts;
            }
        }

        string _Email;
        public string Email
        {
            get
            {
                if (_Email == null)
                {
                    Guid emailGuid;
                    if (isNewOrganization)
                    {
                        emailGuid = new Guid("3c081d6a-b372-4d15-871b-2e859f183493");
                    }
                    else { emailGuid = new Guid("0f4a1439-7f2c-495d-940a-bb41937b0c23"); }
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
                    Guid nameGuid;
                    if (isNewOrganization)
                    {
                        nameGuid = new Guid("38a5b8fd-9d35-43bf-b4d1-8ddb256a6805");
                    }
                    else
                    {
                        nameGuid = new Guid("e434c150-94e4-4cb6-a1de-9fc8b6408380");
                    }
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

        public object ID { get { return this.ReferenceObject.SystemFields.Id; } }

        internal void Delete()
        {
            this.ReferenceObject.Delete();
        }

        public override string ToString()
        {
            if (string.IsNullOrEmpty(Name))
                return base.ToString();
            else return Name;
        }
    }

    public class Contact : IMailRecipient
    {

        //public Contact(ReferenceObject contact_ro)
        //{
        //    this.ReferenceObject = contact_ro;
        //}

        public Contact(string fullName, string eMailString)
        {
            this._FullName = fullName;
            this._Email = eMailString;
        }

        //public ReferenceObject ReferenceObject { get; private set; }

        string _Email;
        public string Email
        {
            get
            {
                //if (_Email == null)
                //    _Email = this.ReferenceObject[new Guid("47bec745-0fc4-4fc8-b20d-0c5e29cb8dab")].GetString();
                return _Email;
            }
        }

        string _FullName;
        public string FullName
        {
            get
            {
                //if (_FullName == null)
                //    _FullName = this.ReferenceObject[new Guid("0612104e-7a10-4cdb-8090-5e61a87caa94")].GetString();
                return _FullName;
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

    public class FactoryContact
    {
        public static Contact CreateContact(ReferenceObject referenceObject)
        {
            string fullName;
            string eMailString;

            if (referenceObject.Reference.ParameterGroup.ReferenceInfo.Guid.ToString() == "c012f7db-2324-4c40-84d6-ca9502021d19")
            {
                fullName = referenceObject[new Guid("730568aa-ac67-4a51-bd8c-8f40068326df")].GetString();
                eMailString = referenceObject[new Guid("fb1f3a22-2cf1-405a-9db8-662ecebf6a45")].GetString();
            }
            else
            {
                fullName = referenceObject[new Guid("0612104e-7a10-4cdb-8090-5e61a87caa94")].GetString();
                eMailString = referenceObject[new Guid("47bec745-0fc4-4fc8-b20d-0c5e29cb8dab")].GetString();
            }

            return new Cancelaria.Contact(fullName, eMailString);
        }
    }

    class RegistryJournal
    {
        public RegistryJournal(ReferenceObject referenceObject)
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

    public class BiblORDRecord
    {
        private static Guid BORD_reference_BiblORD_Guid = new Guid("e718cbfa-a6d1-4c18-b40d-3d1eabaa4d86");    //Guid справочника - "Система учета и контроля над исполнением ОРД"
        private static Guid BORD_class_ORD_GUID = new Guid("f1766d22-3f6b-49ff-b1d9-08d6d48f5909"); //Guid типа - "ОРД"
        private static Guid BORD_param_Name_GUID = new Guid("8f51f04d-091b-4053-81be-e6407d78ef07"); //параметр "Наименование"
        private static Guid BORD_param_ReleaseDate_GUID = new Guid("3b91c892-9f0b-4dd5-a704-9ea5a2129a5e"); //параметр "Дата выпуска документа"
        private static Guid BORD_param_ShowToAll_GUID = new Guid("372bfd15-4715-4f30-95b2-c00ba9c51935"); //параметр "Просмотр всем"
        private static Guid BORD_link_RegistryCancelariaCard_GUID = new Guid("776b674c-4944-460a-a8e0-729efa58911e"); //cвязь 1:1 Регистрационно-контрольная карта Справочника Канцелярия
        private static Guid BORD_link_MailRecipients_GUID = new Guid("d927ffa1-9610-4c6d-90fc-1af3f635ab7c"); //cвязь N:N Адресанты ОРД
        private static Guid BORD_link_DocumentFile_GUID = new Guid("b442573d-602a-4b60-8a11-cbc36790fa88"); //cвязь 1:1 ОРД-Файлы
        private static Guid BORD_item_AccessLevel1Users_GUID = new Guid("4d868da4-4bcf-4a9c-9213-b35c120a8bca"); //объект Уровень доступа 2 справочника Группы и пользователи


        ReferenceObject Original;

        public ReferenceObject GetNewBORDReferenceObject()
        {
            Original = CreateNewBiblORDInDB();
            FillCommonParametersToNewRO();
            FillAccessParametersToNewRO();
            AddFileDocumentToNewRO();
            //MessageBox.Show("Сохраняем в БД");
            //Original.ApplyChanges();
            return Original;
        }

        public void LinkWithRCC(ReferenceObject rc)
        {
            Original.SetLinkedObject(BORD_link_RegistryCancelariaCard_GUID, rc);
        }

        void AddFileDocumentToNewRO()
        {
            Original.SetLinkedObject(BORD_link_DocumentFile_GUID, DocumentFile);
        }

        void FillCommonParametersToNewRO()
        {
            Original[BORD_param_ReleaseDate_GUID].Value = DateTime.Today;
            Original[BORD_param_Name_GUID].Value = Name;
        }

        void FillAccessParametersToNewRO()
        {
            //добавляем непосредственных участников в список рассылки
            foreach (var item in UsersAccess)
                Original.AddLinkedObject(BORD_link_MailRecipients_GUID, item);
            //добавляем в список рассылки согласно уровню доступа
            if (AccessLevel == 0) Original[BORD_param_ShowToAll_GUID].Value = true;
            if (AccessLevel == 1)
            {
                Original[BORD_param_ShowToAll_GUID].Value = false;
                ReferenceObject accessLevel2UserGroup = (new UserReference(Connection)).Find(BORD_item_AccessLevel1Users_GUID);
                Original.AddLinkedObject(BORD_link_MailRecipients_GUID, accessLevel2UserGroup);
            }
            if (AccessLevel == 2)
            {
                Original[BORD_param_ShowToAll_GUID].Value = false;
            }

        }

        int AccessLevel;
        public void SetAccessLevel(int accessLevel)
        {
            AccessLevel = accessLevel;
        }

        public void AddToAccessList(IEnumerable<ReferenceObject> users)
        {
            UsersAccess = users;
        }

        ReferenceObject DocumentFile;
        public void SetFileDocument(ReferenceObject file)
        {
            DocumentFile = file;
        }

        string Name;
        public void SetName(string name)
        { Name = name; }

        ReferenceObject CreateNewBiblORDInDB()
        {
            // получение описания справочника
            ReferenceInfo referenceInfo = Connection.ReferenceCatalog.Find(BORD_reference_BiblORD_Guid);
            // создание объекта для работы с данными
            Reference reference = referenceInfo.CreateReference();
            // поиск типа
            ClassObject classObject = reference.Classes.Find(BORD_class_ORD_GUID);
            // создание
            ReferenceObject newObject = reference.CreateReferenceObject(classObject);

            return newObject;
        }

        IEnumerable<ReferenceObject> UsersAccess;

        public static ServerConnection Connection
        {
            get { return ServerGateway.Connection; }
        }

    }

    /// <summary>
    /// Справочники
    /// ver 1.2
    /// </summary>
    public class References
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

                return _WorkChangeRequestReference;
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

        //private static ClassObject _Class_ProjectManagementWork;
        /// <summary>
        /// тип Работа - справочника Управление проектами
        /// </summary>
        //public static ClassObject Class_ProjectManagementWork
        //{
        //    get
        //    {
        //        if (_Class_ProjectManagementWork == null)
        //            _Class_ProjectManagementWork = ProjectManagementReference.Classes.Find(ProjectManagementWork.PM_class_Work_Guid);

        //        return _Class_ProjectManagementWork;
        //    }
        //}

        /// <summary>
        /// Справочник "Управление проектами"
        /// </summary>
        public static Reference ProjectManagementReference
        {
            get
            {
                if (_ProjectManagementReference == null)
                    _ProjectManagementReference = GetReference(ref _ProjectManagementReference, ProjectManagementReferenceInfo);

                return _ProjectManagementReference;
            }
        }



        /// <summary>
        /// Справочник "Управление проектами"
        /// </summary>
        private static ReferenceInfo ProjectManagementReferenceInfo
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

        // Справочник "Ресурсы"
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

                return _ResourcesReference;
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
        #endregion

        #region "Регистрационно-Контрольные карточки (Канцелярия)"

        static readonly Guid ref_RegistryControlCards_Guid = new Guid("80831dc7-9740-437c-b461-8151c9f85e59");
        private static Reference _RegistryControlCardsReference;
        /// <summary>
        /// Справочник "Регистрационно-Контрольные карточки"
        /// </summary>
        public static Reference RegistryControlCards
        {
            get
            {
                if (_RegistryControlCardsReference == null)
                    _RegistryControlCardsReference = GetReference(ref _RegistryControlCardsReference, RegistryControlCardsReferenceInfo);

                _RegistryControlCardsReference.Refresh();
                return _RegistryControlCardsReference;
            }
        }

        private static ReferenceInfo _RegistryControlCardsReferenceInfo;
        /// <summary>
        /// Справочник "Регистрационно-Контрольные карточки"
        /// </summary>
        public static ReferenceInfo RegistryControlCardsReferenceInfo
        {
            get { return GetReferenceInfo(ref _RegistryControlCardsReferenceInfo, ref_RegistryControlCards_Guid); }
        }
        #endregion

        #region "Статистика загрузки ресурсов"

        // Справочник "Статистика загрузки ресурсов"
        static readonly Guid ref_StatisticConsumptionResource_Guid = new Guid("fb63e9e9-71b3-4afc-aed5-8312a732a60a");        //Guid справочника - "Статистика загрузки ресурсов"
        private static Reference _StatisticConsumptionResourceReference;
        /// <summary>
        /// Справочник "Статистика загрузки ресурсов"
        /// </summary>
        public static Reference StatisticConsumptionResourceReference
        {
            get
            {
                if (_StatisticConsumptionResourceReference == null)
                    _StatisticConsumptionResourceReference = GetReference(ref _StatisticConsumptionResourceReference, StatisticConsumptionResourceReferenceInfo);

                return _StatisticConsumptionResourceReference;
            }
        }

        private static ReferenceInfo _StatisticConsumptionResourceReferenceInfo;
        /// <summary>
        /// Справочник "Статистика загрузки ресурсов"
        /// </summary>
        private static ReferenceInfo StatisticConsumptionResourceReferenceInfo
        {
            get { return GetReferenceInfo(ref _StatisticConsumptionResourceReferenceInfo, ref_StatisticConsumptionResource_Guid); }
        }


        static readonly Guid SCR_class_StatisticConsumptionResource_Guid = new Guid("41606f1a-16c5-4df4-b507-360590ebfe95"); //Guid типа "Загрузка ресурса"
        static ClassObject _Class_StatisticConsumptionResource;
        /// <summary>
        /// класс Нерасходуемые ресурсы
        /// </summary>
        public static ClassObject Class_StatisticConsumptionResource
        {
            get
            {
                if (_Class_StatisticConsumptionResource == null)
                    _Class_StatisticConsumptionResource = StatisticConsumptionResourceReference.Classes.Find(SCR_class_StatisticConsumptionResource_Guid);

                return _Class_StatisticConsumptionResource;
            }
        }
        #endregion

        private static ReferenceInfo GetReferenceInfo(ref ReferenceInfo referenceInfo, Guid referenceGuid)
        {
            if (referenceInfo == null)
                referenceInfo = Connection.ReferenceCatalog.Find(referenceGuid);

            return referenceInfo;
        }
    }

    //// <summary>
    /// HTML письмо
    /// ver 2.2 (с применением интерфейса)
    /// </summary>
    class HTML_MailMessageI
    {
        public HTML_MailMessageI()
        {
            Subject = "";
            BodyText = "";
            MailAttachments = new List<ReferenceObject>();
            MailRecipients = new List<IMailRecipient>();
            MailCopyRecipients = new List<IMailRecipient>();
        }

        public HTML_MailMessageI(string messageSubject, string messageText, List<IMailRecipient> mailRecipients, List<ReferenceObject> works)
        {
            this.Subject = messageSubject;
            this.BodyText = messageText;
            this.MailRecipients = mailRecipients;
            this.MailAttachments = works;
            this.MailCopyRecipients = new List<IMailRecipient>();
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
        public List<IMailRecipient> MailRecipients { get; set; }
        //List<IMailRecipient> _MailCopyRecipients;
        public List<IMailRecipient> MailCopyRecipients { get; set; }
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

        public static bool IsValidEmail(string email)
        {
            //regular expression pattern for valid email
            //addresses, allows for the following domains:
            //com,edu,info,gov,int,mil,net,org,biz,name,museum,coop,aero,pro,tv
            string pattern = @"^[-a-zA-Z0-9][-._a-zA-Z0-9]*@[-.a-zA-Z0-9]+(\.[-.a-zA-Z0-9]+)*\.
    (com|edu|info|gov|int|mil|net|org|biz|name|museum|coop|aero|pro|tv|[a-zA-Z]{2})$";
            //Regular expression object
            Regex check = new Regex(pattern, RegexOptions.IgnorePatternWhitespace);
            //boolean variable to return to calling method
            bool valid = false;

            //make sure an email address was provided
            if (string.IsNullOrEmpty(email))
            {
                valid = false;
            }
            else
            {
                //use IsMatch to validate the address
                valid = check.IsMatch(email);
            }
            //return the value to the calling method
            return valid;
        }

#if !server
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
            List<string> emails = new List<string>();
            foreach (IMailRecipient recipient in MailRecipients)
            {
                if (recipient.UserReferenceObject != null)
                {
                    UserReferenceObject recipient_uro = recipient.UserReferenceObject as UserReferenceObject;
                    if (recipient_uro is User)
                    {
                        tempUsers.Add(recipient_uro);
                    }
                    else if (recipient_uro is UsersGroup)
                    {
                        List<User> internUsers = recipient_uro.GetAllInternalUsers();
                        if (internUsers != null) tempUsers.AddRange(recipient_uro.GetAllInternalUsers());
                    }
                    else continue;
                }
                else if (!string.IsNullOrWhiteSpace(recipient.Email))
                {
                    if (emails.IndexOf(recipient.Email) >= 0) continue;
                    else
                    {
                        MailMessage.To.Add(new EMailAddress(recipient.Email));
                        emails.Add(recipient.Email);
                    }
                }
                else throw new Exception("Отсутствуют данные адресата для отправки письма!");
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
            foreach (IMailRecipient recipient in MailCopyRecipients)
            {
                if (recipient.UserReferenceObject != null)
                {
                    UserReferenceObject recipient_uro = recipient.UserReferenceObject as UserReferenceObject;
                    if (recipient_uro is User)
                    {
                        tempUsers.Add(recipient_uro);
                    }
                    else if (recipient_uro is UsersGroup)
                    {
                        List<User> internUsers = recipient_uro.GetAllInternalUsers();
                        if (internUsers != null) tempUsers.AddRange(recipient_uro.GetAllInternalUsers());
                    }
                    else continue;
                }
                else if (!string.IsNullOrWhiteSpace(recipient.Email))
                {
                    MailMessage.To.Add(new EMailAddress(recipient.Email));
                }
                else throw new Exception("Отсутствуют данные адресата для отправки письма!");
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
                bodyFooter = "";

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

    interface IMailRecipient
    {
        UserReferenceObject UserReferenceObject { get; }
        string Email { get; }
    }

    public class MailRecipient : IMailRecipient
    {

        public MailRecipient(UserReferenceObject user)
        {
            this.UserReferenceObject = user;
            string email = null;
            try
            {
                email = this.UserReferenceObject[new Guid("e17261a6-3082-4d2a-8fde-4aee7156ee80")].GetString();
            }
            catch { }
            this.Email = email;
            this.Name = user.ToString();
        }
        public MailRecipient(string name, string email)
        {
            this.Name = name;
            this.Email = email;
        }

        public string Email
        {
            get;
            private set;
        }

        public UserReferenceObject UserReferenceObject
        {
            get;
            private set;
        }

        public string Name { get; set; }
    }

}