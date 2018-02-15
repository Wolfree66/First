//#define server

//#define test

//!!!!!!!!!!!!!!!!!!!!!!!!!!!!
//Не забудь переключить ГУИД инструкции с тестового на боевой

/*
                    Ссылки:
                    TFlex.DOCs.Model.Processes.dll
                    TFlex.DOCs.UI.Common.dll
                    TFlex.DOCs.UI.Types.dll
                    TFlex.DOCs.UI.Controls.dll
                    TFlex.DOCs.UI.Objects.dll
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.Mail;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Stages;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.References.GlobalParameters;
using System.IO;

#if !server
using TFlex.DOCs.UI.Objects.Managers;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros.ObjectModel;
#endif

using TFlex.DOCs.Model.References.Calendar;


namespace ErrandControl
{


    public class Macro : MacroProvider
    {
        public Macro(MacroContext context)
            : base(context)
        {
        }

        //Справочник - "Контроль поручений"
        private static readonly Guid param_ErrandDocNum_GUID = new Guid("f89f8ad5-7c33-4d68-8428-105b5991a3e9");//Guid параметра "Документ с поручением"
        private readonly Guid param_Category_GUID = new Guid("503cbb52-0204-4feb-9a17-cc83c5d4395e");//Guid параметра "Категория"

        //Даты

        private readonly Guid param_ErrandEndDate_GUID = new Guid("24bf25f7-bf81-405f-8f4e-75dda56e7ed3");//Guid параметра "Плановая дата выполнения поручения"
        private readonly Guid param_VerificationDate_GUID = new Guid("07309e1e-4e97-492f-a2ad-303616988c40");//Guid параметра "Дата верификации"

        //результаты
        private readonly Guid param_PerformerResult_GUID = new Guid("22124188-0f75-4417-961d-36a66cac0e7a");//Guid параметра "Результат исполнения"
        private readonly Guid param_VerificationResult_GUID = new Guid("8632c88c-fc67-4f1e-a650-72d11a35f2e0");//Guid параметра "Результат верификации"
        private readonly Guid param_ValidationResult_GUID = new Guid("7fc12479-9f81-44b5-a6f7-b74706c52b66");//Guid параметра "Результат валидации"

        //роли и должности
        private readonly Guid link_Director_GUID = new Guid("587b07dc-fc72-43f8-9fb2-e698f52db3c3");//Guid связи n:1 "Руководитель"

        private readonly Guid param_ExtErrand_GUID = new Guid("5ed4ffce-30c7-4308-8a8c-1e2974af379e");//Guid параметра "Внешнее поручение"
                                                                                                      //private readonly Guid param_ExternalPerformerEmail_GUID = new Guid("6131602b-933a-4895-8aeb-542c216cfa1c");//Guid параметра "E-mail внешнего исполнителя"
                                                                                                      //private readonly Guid param_ExtPerformer_GUID = new Guid("52129989-7ec0-4984-9c1a-77658ae1e763");//Guid параметра "Внешний исполнитель"

        private readonly Guid link_Performer_GUID = new Guid("7cd90bfa-e37c-4265-a7e1-f65812940575");//Guid связи n:1 "Исполнитель поручения"
        private readonly Guid link_Checker_GUID = new Guid("10c6727f-f6a7-4550-a0ee-da862ff14d12");//Guid связи n:1 "Контролёр поручения"
                                                                                                   //private readonly Guid link_PositionDirector_GUID = new Guid("830bdf4f-6e8c-4c83-9790-be29b3fe55be");//Guid связи n:1 "Должность Руководителя"
        private readonly Guid link_PositionPerformer_GUID = new Guid("8c18ac88-697a-4d22-afdc-49dd557b5cca");//Guid связи n:1 "Должность Исполнителя"
                                                                                                             //private readonly Guid link_PositionChecker_GUID = new Guid("3508946b-d8d2-407c-8fe1-1205f158e794");//Guid связи n:1 "Должность Контролёра"

        private readonly Guid link_DepartmentPerformer_GUID = new Guid("a93bcf81-eb0b-4856-a905-8147e28c54e6");//Guid связи n:1 "Подразделение исполнителя"

        private readonly Guid link_ToInputData_GUID = new Guid("5b41779b-122e-4a53-8feb-a6388e44c81d");//Guid связи n:n "Исходные данные"

        //Guid справочника - "Процедуры"
        //private readonly Guid reference_Procedures_GUID = new Guid("61d922d0-4d60-49f1-a6a0-9b6d22d3d8b3");
        //Guid связи 1:n "Процессы"
        //private readonly Guid link_ToProcesses_GUID = new Guid("44c3a825-ef4a-4ecb-a291-1048db42d5cb");

        //Guid связи n:1 "Ответственный - Группы и пользователи" справочника РКК
        private readonly Guid link_ToResponsible_GUID = new Guid("6a69454d-7803-4cd8-8874-cde190185b7b");
        //Guid параметра "Регистрационный номер" справочника РКК
        private readonly Guid param_RegNum_GUID = new Guid("ac452c8a-b17a-4753-be3f-195d76480d52");
        //Guid связи n:n "Документы" справочника РКК
        private readonly Guid link_ToDocuments_GUID = new Guid("aab1cd65-fe2d-4fc0-8db2-4ff2b3d215d5");
        //связь со справочником Группы и пользователи - Кому
        private Guid guid_komu = new Guid("a3746fbd-7bda-4326-8e4b-f6bbefc2ecce");

        //Guid файла "Контроль поручений_Исполнитель.pdf" - справочника "Файлы" на тестовом сервере
        //private readonly Guid file_PerformShortInstr_GUID = new Guid("4c5b2e7c-fd32-414f-abf4-ded714ce665e");
        //Guid файла "Контроль поручений_Исполнитель.pdf" - справочника "Файлы" на боевом сервере
        private readonly Guid file_PerformShortInstr_GUID = new Guid("5f49d507-fb70-4318-845f-ebdaee7cea77");

        public override void Run()
        {
            // MessageBox.Show("Процесс контроля поручений - запущен.");
        }
#if !server
        public void TestClient()
        {
            ReferenceObject errand_ro = Context.ReferenceObject;
            Errand newErrand = new Errand(errand_ro);
            CorrectionRequest cr = new CorrectionRequest(newErrand);
            cr.ShowDialog();
        }

        public void Test()
        {
            string req = @"***** 13.12.2017 10:54:45; Тестов1 Тест1 Тестович1; Поручение корректируется Руководителем. *****



***** 13.12.2017 10:52:49; Тестов2 Тест2 Тестович2; Требуется коррекция. *****

Причина: причина
Изменение параметра: Плановая дата выполнения - 11.01.2018


***** 13.12.2017 10:47:48; Тестов3 Тест3 Тестович3; Промежуточная запись. *****
запись


***** 13.12.2017 10:47:19; Тестов3 Тест3 Тестович3; C поручением ознакомлен. *****



***** 13.12.2017 10:47:03; Тестов1 Тест1 Тестович1; Отправлено новое поручение *****

<Текст поручения>: Поручение 2
<Документ с поручением>: документ
<Срок исполнения>: 12.01.2018
<Ожидаемый результат>: результ
<Руководитель>: Тестов1 Тест1 Тестович1
<Исполнитель>: Тестов3 Тест3 Тестович3
<Контролёр>: Тестов2 Тест2 Тестович2


";
            string errandText;
            string errandDocumentForErrand;
            string errandPlanEndDate;
            string errandProposedResult;
            string errandDirector;
            string errandPerformer;
            string errandChecker;
            //ищем в логе последнюю запись о реквизитах поручения
            //GetLastParametersFromLog(req, out errandText, out errandDocumentForErrand, out errandPlanEndDate,
            //out errandProposedResult, out errandDirector, out errandPerformer, out errandChecker);
            //ReferenceObject errand_ro = ErrandControlReference.Reference.Find(2249);
            //Errand newErrand = new Errand(errand_ro);
            //CancelCorrection(newErrand);
        }
#endif

        #region Элементы бизнес логики

        internal bool StartControlProcess(Errand errand)
        {
            if (errand.Stage != Errand.ErrandStages.New)
            {
                System.Windows.Forms.MessageBox.Show("Поручение не находится в стадии - \"Новое поручение\"");
                return false;
            }
            //проверяем заполненность необходимых полей
            if (!errand.IsRequiredRequisitesFilled) return false;

            //информируем участников процесса
            string textSubj = "Новое поручение №" + errand.ErrandNumber.ToString() + ". " + errand.DocumentForErrand.ToString() + ". Срок: " + errand.PlanEndDate.ToShortDateString();
            string textMsg = "В АСУ \"Динамика\", в модуле \"Контроль поручений\" - " + CurrentUserName
                + " создал новое поручение. Ознакомьтесь с поручением, перейдя ко вложенному объекту.";

            List<UserReferenceObject> mailRecipients = new List<UserReferenceObject>() { errand.Checker };
            string email = "";
            if (!errand.IsPerformerExternal) mailRecipients.Add(errand.Performer);
            else
                email = errand.ExternalPerformerEMail;
            SendHTMLMessage(textSubj, textMsg, mailRecipients, email, new List<Errand>() { errand });

            //присваиваем дату отправки
            DateTime sendDate = DateTime.Now;
            //записываем в лог
            try
            {
                string comment = LogRequisiteComment(errand);
                if (!AddRecordToLog(errand, sendDate, "Отправлено новое поручение", (int)Errand.Role.None, false, comment)) return false;
            }
            catch (Errand.Exception_CancelAllActivities e)
            { return false; }
            errand.SendToPerformerDate = sendDate;
            //если Исполнитель внешний то присваиваем результат выполнения = Не выполнено
            if (errand.IsPerformerExternal)
            {
                errand.PerformerResult = Errand.PerformerResultValue.NotDone;
            }
            else errand.PerformerResult = Errand.PerformerResultValue.NotAccepted;

            errand.SaveSecondaryParametersToDB();
            //изменяем стадию на "В работе"
            errand.ChangeStageTo(Errand.ErrandStages.InWork);
            return true;
        }

        void StartCorrection(Errand errand)
        {
            if (errand.Stage == Errand.ErrandStages.Closed)
            {
                System.Windows.Forms.MessageBox.Show("Поручение находится в стадии - \"Закрыто\", изменение невозможно.");
                return;
            }
            DateTime sendDate = DateTime.Now;
            AddRecordToLog(errand, sendDate, "Поручение корректируется Руководителем.", Errand.Role.Director, false);
            errand.ValidationResult = Errand.ValidationResultValue.Correction;
            errand.SaveSecondaryParametersToDB();
        }


        internal void EndTheCorrection(Errand errand)
        {
            if (errand.Stage == Errand.ErrandStages.Closed)
            {
                System.Windows.Forms.MessageBox.Show("Поручение находится в стадии - \"Закрыто\", изменение невозможно.");
                return;
            }
            //проверяем заполненность необходимых полей
            if (!errand.IsRequiredRequisitesFilled) { return; }



            //присваиваем дату отправки
            DateTime sendDate = DateTime.Now;
            //записываем в лог
            try
            {
                string comment = LogRequisiteComment(errand);
                if (!AddRecordToLog(errand, sendDate, "Поручение скорректировано.", Errand.Role.Director, true, comment)) return;
            }
            catch (Errand.Exception_CancelAllActivities e)
            { return; }
            errand.SendToPerformerDate = sendDate;

            if (errand.IsPerformerExternal)
            {
                errand.PerformerResult = Errand.PerformerResultValue.NotDone;
            }
            else errand.PerformerResult = Errand.PerformerResultValue.NotAccepted;

            errand.VerificationResult = Errand.VerificationResultValue.NotVerified;
            errand.ValidationResult = Errand.ValidationResultValue.NotValidated;
            //если Исполнитель внешний то присваиваем результат выполнения = Не выполнено
            if (errand.IsPerformerExternal)
            {
                errand.PerformerResult = Errand.PerformerResultValue.NotDone;
            }
            else errand.PerformerResult = Errand.PerformerResultValue.NotAccepted;
            errand.PerformerDoneDate = null;

            errand.VerificationResult = Errand.VerificationResultValue.NotVerified;
            errand.VerificationDate = null;

            errand.ValidationResult = Errand.ValidationResultValue.NotValidated;
            errand.ValidationDate = null;
            if (errand.ActualCorrectionRequest != null)
                errand.ActualCorrectionRequest.Apply("");
            errand.SaveSecondaryParametersToDB();
            //изменяем стадию на "В работе"
            errand.ChangeStageTo(Errand.ErrandStages.InWork);

            string textSubj = "Скорректированное поручение №" + errand.ErrandNumber.ToString() + ". " + errand.DocumentForErrand.ToString() + ". Срок: " + errand.PlanEndDate.ToShortDateString();
            string textMsg = string.Format(@"В АСУ ""Динамика"", в модуле ""Контроль поручений"" -
                {0} скорректировал поручение. Комментарий: {1}. Ознакомьтесь с поручением, перейдя ко вложенному объекту.", CurrentUserName, errand.DirectorComment);

            //информируем участников процесса
            InformPerformerAndChecker(errand, textSubj, textMsg);
        }

#if !server
        void CancelCorrection(Errand errand)
        {

            string errandText;
            string documentForErrand;
            string planEndDate;
            string proposedResult;
            string directorName;
            int coefficientOfImportance = 0;
            string performerName;
            bool isExternalPerformer;
            string externalPerformerEMail;

            string checkerName;
            //ищем в логе последнюю запись о реквизитах поручения
            bool isGetParams = GetLastParametersFromLog(errand.Log, out errandText, out documentForErrand, out planEndDate, out coefficientOfImportance,
            out proposedResult, out directorName, out performerName, out isExternalPerformer, out externalPerformerEMail, out checkerName);
            if (!isGetParams)
            {
                MessageBox.Show("Невозможно определить исходные параметры, используйте команду Завершить коррекцию!", "Невозможно отменить коррекцию!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                //сверяем с текущими параметрами
                StringBuilder errors = new StringBuilder();
                if (errand.Text.Trim(' ') != errandText) errand.Text = errandText;
                if (errand.DocumentForErrand.Trim(' ') != documentForErrand) errand.DocumentForErrand = documentForErrand;
                if (errand.PlanEndDate.ToShortDateString() != planEndDate) errand.PlanEndDate = DateTime.Parse(planEndDate);
                if (errand.ProposedResult.Trim(' ') != proposedResult) errand.ProposedResult = proposedResult;
                if (errand.Director == null || errand.Director.FullName.ToString().Trim(' ') != directorName) errand.Director = UserReference.FindUser(directorName);
                if (errand.Checker == null || errand.Checker.FullName.ToString().Trim(' ') != checkerName) errand.Checker = UserReference.FindUser(checkerName);
                errand.CoefficientOfImportance = coefficientOfImportance;
                errand.IsPerformerExternal = isExternalPerformer;
                if (isExternalPerformer)
                {
                    errand.ExternalPerformer = performerName;
                    errand.ExternalPerformerEMail = externalPerformerEMail;
                }
                else
                {
                    if (errand.Checker == null || errand.Checker.FullName.ToString().Trim(' ') != checkerName) errand.Checker = UserReference.FindUser(checkerName);
                }

                errand.SaveRequisitesToDB();
                //если совпадает то меняем результат валидации на Не валидировано
                if (string.IsNullOrWhiteSpace(errors.ToString()))
                {
                    errand.ValidationResult = Errand.ValidationResultValue.NotValidated;
                    errand.ValidationDate = null;
                    errand.SaveSecondaryParametersToDB();
                    DateTime actionDate = DateTime.Now;
                    bool recordAdded = AddRecordToLog(errand, actionDate, "Руководитель отменил коррекцию.", Errand.Role.Director, false);
                    errand.SaveSecondaryParametersToDB();
                    if (!recordAdded) MessageBox.Show("Не удалось изменить log", "Не удалось изменить реквизиты!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                else //если вернуть не получилось выдаём информацию о том что не получилось
                {
                    MessageBox.Show("Не удалось изменить следующие реквизиты, (попробуйте изменить их вручную):\n" + errors.ToString(), "Не удалось изменить реквизиты!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }


#endif

        internal void AutomaticChangePlanDateByCorrectionRequest(Errand errand)
        {
            if (errand.Stage == Errand.ErrandStages.Closed)
            {
                System.Windows.Forms.MessageBox.Show("Поручение находится в стадии - \"Закрыто\", изменение невозможно.");
                return;
            }
            //проверяем заполненность необходимых полей
            if (!errand.IsRequiredRequisitesFilled) { return; }
            errand.ValidationResult = Errand.ValidationResultValue.Correction;
            errand.SaveSecondaryParametersToDB();

            errand.PlanEndDate = errand.ActualCorrectionRequest.RequisiteChangeNotices.Where(rcn => rcn.ChangingParameterName == Errand.Requisites.PlanDate).FirstOrDefault().NewValue;
            //присваиваем дату отправки
            DateTime sendDate = DateTime.Now;
            //записываем в лог
            try
            {
                string comment = LogRequisiteComment(errand);
                if (!AddRecordToLog(errand, sendDate, "Поручение скорректировано.", Errand.Role.None, false, comment)) return;
            }
            catch (Errand.Exception_CancelAllActivities e)
            { return; }
            errand.SendToPerformerDate = sendDate;

            if (errand.IsPerformerExternal)
            {
                errand.PerformerResult = Errand.PerformerResultValue.NotDone;
            }
            else errand.PerformerResult = Errand.PerformerResultValue.NotAccepted;

            errand.VerificationResult = Errand.VerificationResultValue.NotVerified;
            errand.ValidationResult = Errand.ValidationResultValue.NotValidated;
            //если Исполнитель внешний то присваиваем результат выполнения = Не выполнено
            if (errand.IsPerformerExternal)
            {
                errand.PerformerResult = Errand.PerformerResultValue.NotDone;
            }
            else errand.PerformerResult = Errand.PerformerResultValue.NotAccepted;
            errand.PerformerDoneDate = null;

            errand.VerificationResult = Errand.VerificationResultValue.NotVerified;
            errand.VerificationDate = null;

            errand.ValidationResult = Errand.ValidationResultValue.NotValidated;
            errand.ValidationDate = null;
            if (errand.ActualCorrectionRequest != null)
                errand.ActualCorrectionRequest.Apply("Автоматическое изменение плановой даты выполнения, по истечению срока давности.");
            errand.SaveRequisitesToDB();
            errand.SaveSecondaryParametersToDB();
            //изменяем стадию на "В работе"
            errand.ChangeStageTo(Errand.ErrandStages.InWork);
            /*
            string textSubj = "Скорректированное поручение №" + ErrandNumber.ToString() + ". " + DocumentForErrand.ToString() + ". Срок: " + PlanEndDate.ToShortDateString();
            string textMsg = string.Format(@"В АСУ ""Динамика"", в модуле ""Контроль поручений"" -
                {0} скорректировал поручение. Комментарий: {1}. Ознакомьтесь с поручением, перейдя ко вложенному объекту.", CurrentUserName, DirectorComment);

            //информируем участников процесса
            InformPerformerAndChecker(textSubj, textMsg);
             */
        }

        void ValidateWithResult(Errand errand, string validationResult)
        {
            if (errand.Stage == Errand.ErrandStages.Closed)
            {
                System.Windows.Forms.MessageBox.Show("Поручение находится в стадии - \"Закрыто\", изменение невозможно.");
                return;
            }
            if (validationResult == Errand.ValidationResultValue.NotDone)
            {

                //проверяем заполненность необходимых полей
                if (!errand.IsRequiredRequisitesFilled) { return; }

                //присваиваем дату отправки
                DateTime sendDate = DateTime.Now;
                //записываем в лог
                try
                {
                    string comment = LogRequisiteComment(errand);
                    if (!AddRecordToLog(errand, sendDate, "Поручение не выполнено.", Errand.Role.Director, true, comment)) return;
                }
                catch (Errand.Exception_CancelAllActivities e)
                { return; }
                errand.SendToPerformerDate = sendDate;

                if (errand.IsPerformerExternal)
                {
                    errand.PerformerResult = Errand.PerformerResultValue.NotDone;
                }
                else errand.PerformerResult = Errand.PerformerResultValue.NotAccepted;

                errand.VerificationResult = Errand.VerificationResultValue.NotVerified;
                errand.ValidationResult = Errand.ValidationResultValue.NotValidated;
                //если Исполнитель внешний то присваиваем результат выполнения = Не выполнено
                if (errand.IsPerformerExternal)
                {
                    errand.PerformerResult = Errand.PerformerResultValue.NotDone;
                }
                else errand.PerformerResult = Errand.PerformerResultValue.NotAccepted;
                errand.PerformerDoneDate = null;

                errand.VerificationResult = Errand.VerificationResultValue.NotVerified;
                errand.VerificationDate = null;

                errand.ValidationResult = Errand.ValidationResultValue.NotValidated;
                errand.ValidationDate = null;

                errand.SaveSecondaryParametersToDB();
                string textSubj = "Не выполненное поручение №" + errand.ErrandNumber.ToString() + ". " + errand.DocumentForErrand.ToString() + ". Срок: " + errand.PlanEndDate.ToShortDateString();
                string textMsg = string.Format(@"В АСУ ""Динамика"", в модуле ""Контроль поручений"" -
                {0} отметил поручение как  невыполненное. Комментарий: {1}. Ознакомьтесь с поручением, перейдя ко вложенному объекту.", CurrentUserName, errand.DirectorComment);

                //информируем участников процесса
                InformPerformerAndChecker(errand, textSubj, textMsg);
            }
            else
            {
                DateTime date = DateTime.Now;
                //записываем в лог
                try
                {
                    string actionName = string.Format("Поручение {0}", validationResult.ToLower());

                    if (!AddRecordToLog(errand, date, actionName, Errand.Role.Director, true)) return;
                }
                catch (Errand.Exception_CancelAllActivities e)
                { return; }

                errand.ValidationResult = validationResult;
                errand.ValidationDate = date;
                errand.SaveSecondaryParametersToDB();

                string textSubj = errand.ValidationResult + ", поручение №" + errand.ErrandNumber.ToString() + ". Документ: " +
                    errand.DocumentForErrand + ". Срок: " + errand.PlanEndDate.ToShortDateString() + ". " + errand.Text;
                string textMsg = "В АСУ \"Динамика\", в модуле \"Контроль поручений\", Руководитель - " + CurrentUserName
                    + ", валидировал поручение №" + errand.ErrandNumber.ToString() + ", с результатом - " + errand.ValidationResult.ToLower()
                    + "<br>Комментарий: " + errand.DirectorComment;
                InformPerformerAndChecker(errand, textSubj, textMsg);

                errand.ChangeStageTo(Errand.ErrandStages.Closed);
            }
        }

        private static string LogRequisiteComment(Errand errand)
        {
            string comment =
                Errand.RequisiteNameForLog.Text + errand.Text.Trim(' ') + Errand.RequisiteNameForLog.Text + "\r\n" +
                Errand.RequisiteNameForLog.DocumentForErrand + errand.DocumentForErrand.Trim(' ') + Errand.RequisiteNameForLog.DocumentForErrand + "\r\n" +
                Errand.RequisiteNameForLog.PlanEndDate + errand.PlanEndDate.ToShortDateString() + Errand.RequisiteNameForLog.PlanEndDate + "\r\n" +
                Errand.RequisiteNameForLog.ProposedResult + errand.ProposedResult.Trim(' ') + Errand.RequisiteNameForLog.ProposedResult + "\r\n" +
                Errand.RequisiteNameForLog.Director + errand.Director.FullName.ToString().Trim(' ') + Errand.RequisiteNameForLog.Director + "\r\n" +
                Errand.RequisiteNameForLog.Checker + errand.Checker.FullName.ToString().Trim(' ') + Errand.RequisiteNameForLog.Checker + "\r\n" +
                Errand.RequisiteNameForLog.ImportanceCoefficient + errand.CoefficientOfImportance.ToString().Trim(' ') + Errand.RequisiteNameForLog.ImportanceCoefficient + "\r\n" +
                Errand.RequisiteNameForLog.IsExternalPerformer + errand.IsPerformerExternal.ToString().Trim(' ') + Errand.RequisiteNameForLog.IsExternalPerformer + "\r\n";
            if (errand.IsPerformerExternal)
            {
                comment += Errand.RequisiteNameForLog.Performer + errand.ExternalPerformer.Trim(' ') + Errand.RequisiteNameForLog.Performer + "\r\n" +
                Errand.RequisiteNameForLog.ExternalPerformerEMail + errand.ExternalPerformerEMail.Trim(' ') + Errand.RequisiteNameForLog.ExternalPerformerEMail + "\r\n";
            }
            else
            {
                comment += Errand.RequisiteNameForLog.Performer + errand.Performer.FullName.ToString().Trim(' ') + Errand.RequisiteNameForLog.Performer + "\r\n";
            }
            return comment;
        }

        private bool GetLastParametersFromLog(string log, out string errandText, out string documentForErrand,
            out string planEndDate, out int coefficientOfImportance, out string proposedResult, out string directorName, out string performerName,
            out bool isExternalPerformer, out string externalPerformerEMail,
            out string checkerName)
        {
            string stringForSearch = Errand.RequisiteNameForLog.Text;
            try
            {
                int startRequisites = log.IndexOf(stringForSearch);

                string requisites = log.Remove(0, startRequisites);
                int endRequisites = requisites.IndexOf("\r\n*****");
                if (endRequisites > 0)
                {
                    requisites = requisites.Remove(endRequisites, requisites.Length - endRequisites);
                }
                errandText = GetValueForRequisite(requisites, Errand.RequisiteNameForLog.Text);
                documentForErrand = GetValueForRequisite(requisites, Errand.RequisiteNameForLog.DocumentForErrand);
                planEndDate = GetValueForRequisite(requisites, Errand.RequisiteNameForLog.PlanEndDate);
                proposedResult = GetValueForRequisite(requisites, Errand.RequisiteNameForLog.ProposedResult);
                directorName = GetValueForRequisite(requisites, Errand.RequisiteNameForLog.Director);
                performerName = GetValueForRequisite(requisites, Errand.RequisiteNameForLog.Performer);
                checkerName = GetValueForRequisite(requisites, Errand.RequisiteNameForLog.Checker);

                string coefImportance = GetValueForRequisite(requisites, Errand.RequisiteNameForLog.ImportanceCoefficient);
                if (!int.TryParse(coefImportance, out coefficientOfImportance))
                { throw new ArgumentOutOfRangeException("Невозможно получить Коэффициент важности, по строке: " + coefImportance); }

                string isExternalPerformerString = GetValueForRequisite(requisites, Errand.RequisiteNameForLog.IsExternalPerformer);
                if (bool.TryParse(isExternalPerformerString, out isExternalPerformer))
                {
                    if (isExternalPerformer)
                    {
                        externalPerformerEMail = GetValueForRequisite(requisites, Errand.RequisiteNameForLog.ExternalPerformerEMail);
                    }
                    else
                        externalPerformerEMail = "";
                }
                else
                { throw new ArgumentOutOfRangeException("Невозможно проверить внешний ли исполнитель, по строке: " + isExternalPerformerString); }
                return true;
            }
            catch (ArgumentOutOfRangeException e)
            {
                errandText = "";
                documentForErrand = "";
                planEndDate = "";
                proposedResult = "";
                directorName = "";
                performerName = "";
                checkerName = "";
                isExternalPerformer = false;
                externalPerformerEMail = "";
                coefficientOfImportance = 0;
                return false;
            }
        }

        private string GetValueForRequisite(string requisitesText, string requisiteName)
        {
            string param = requisitesText.Remove(0, requisitesText.IndexOf(requisiteName) + requisiteName.Length);
            int endIndex = param.IndexOf(requisiteName);
            param = param.Remove(endIndex, param.Length - endIndex);
            return param.Trim(' ');
        }

        void VerificateWithResult(Errand errand, string verificationResult)
        {
            if (errand.Stage == Errand.ErrandStages.Closed)
            {
                System.Windows.Forms.MessageBox.Show("Поручение находится в стадии - \"Закрыто\", изменение невозможно.");
                return;
            }

            if (verificationResult == Errand.VerificationResultValue.DoesNotMatch ||
                verificationResult == Errand.VerificationResultValue.Match)
            {
                DateTime date = DateTime.Now;
                //записываем в лог
                try
                {
                    string actionName = string.Format("Поручение {0}", verificationResult.ToLower());

                    if (!AddRecordToLog(errand, date, actionName, Errand.Role.Checker, true)) return;
                }
                catch (Errand.Exception_CancelAllActivities e)
                { return; }

                errand.VerificationResult = verificationResult;
                errand.VerificationDate = date;

                errand.SaveSecondaryParametersToDB();
            }
            else if (verificationResult == Errand.VerificationResultValue.NeedCorrection)
            {

                if (errand.VerificationResult == Errand.VerificationResultValue.NeedCorrection)
                {
                    if (System.Windows.Forms.MessageBox.Show(("Вы уже информировали Руководителя о необходимости коррекции, хотите уведомить Руководителя повторно?"),
                                                             "Вы нажали кнопку - \"Требуется Коррекция\"", System.Windows.Forms.MessageBoxButtons.OKCancel, System.Windows.Forms.MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.Cancel)
                    { return; }
                }

                DateTime date = DateTime.Now;

                //Создаём запрос коррекции
                CorrectionRequest newCorrection = null;

                if (errand.ActualCorrectionRequest == null)
                {
                    newCorrection = new CorrectionRequest(errand);
                }
                else { newCorrection = errand.ActualCorrectionRequest; }
#if !server
                if (!newCorrection.ApproveByChecker()) { return; }
                else
                {
                    errand.ActualCorrectionRequest = newCorrection;
                    errand.CheckerComment = errand.ActualCorrectionRequest.Reason;
                }
#endif
                //записываем в лог
                try
                {
                    if (!AddRecordToLog(errand, date, "Требуется коррекция.", Errand.Role.Checker, false, errand.ActualCorrectionRequest.ToString())) return;
                }
                catch (Errand.Exception_CancelAllActivities e)
                { return; }
                errand.VerificationResult = verificationResult;
                errand.VerificationDate = date;

                errand.SaveSecondaryParametersToDB();
                string textSubj = "Требуется коррекция. Поручение №" + errand.ErrandNumber.ToString() +
                    ". Документ: " + errand.DocumentForErrand + ". Срок: " + errand.PlanEndDate.ToShortDateString() + ". " + errand.Text;
                string textMsg = "В АСУ \"Динамика\", в модуле \"Контроль поручений\", Контролёр - " + CurrentUserName
                    + ", запросил коррекцию поручения №" + errand.ErrandNumber.ToString() + ". " + "<br>Комментарий: " + errand.CheckerComment;
                InformPerformerAndDirector(errand, textSubj, textMsg);
            }
            else { System.Windows.Forms.MessageBox.Show("Не корректный результат верификации!"); }
        }
#if !server
        void PerformerResultChange(Errand errand, string performerResult)
        {
            if (errand.Stage == Errand.ErrandStages.Closed)
            {
                System.Windows.Forms.MessageBox.Show("Поручение находится в стадии - \"Закрыто\", изменение невозможно.");
                return;
            }

            if (performerResult == Errand.PerformerResultValue.Done)
            {
                if (errand.PerformerResult == performerResult)
                {
                    if (System.Windows.Forms.MessageBox.Show((string.Format("Вы уже отметили это поручение как {0}.\n Вы хотите изменить комментарий и уведомить Руководителя повторно?", performerResult.ToLower())),
                                                             "Вы нажали кнопку - \"Поручение выполнено\"", System.Windows.Forms.MessageBoxButtons.OKCancel, System.Windows.Forms.MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.Cancel)
                    { return; }
                }

                DateTime date = DateTime.Now;
                //записываем в лог
                try
                {
                    string actionName = string.Format("Поручение {0}", performerResult.ToLower());

                    if (!AddRecordToLog(errand, date, actionName, Errand.Role.Performer, true)) return;
                }
                catch (Errand.Exception_CancelAllActivities e)
                { return; }

                errand.PerformerResult = performerResult;
                errand.PerformerDoneDate = date;

                if (errand.VerificationResult == Errand.VerificationResultValue.DoesNotMatch)
                {
                    errand.VerificationResult = Errand.VerificationResultValue.NotVerified;
                    errand.VerificationDate = null;
                }
                errand.SaveSecondaryParametersToDB();

                string textSubj = performerResult + ", поручение №" + errand.ErrandNumber.ToString() + ". Документ: " +
                    errand.DocumentForErrand + ". Срок: " + errand.PlanEndDate.ToShortDateString() + ". " + errand.Text;
                string textMsg = "В АСУ \"Динамика\", в модуле \"Контроль поручений\", Исполнитель - " + CurrentUserName
                    + ", отметил поручение №" + errand.ErrandNumber.ToString() + ", как - " + performerResult.ToLower()
                    + "<br>Комментарий: " + errand.PerformerComment;
                InformChecker(errand, textSubj, textMsg);
            }
            else if (performerResult == Errand.PerformerResultValue.NotAccepted)
            {
                errand.PerformerAcceptDate = null;
                errand.PerformerResult = Errand.PerformerResultValue.NotAccepted;
                errand.SaveSecondaryParametersToDB();
            }
            else if (performerResult == Errand.PerformerResultValue.NeedCorrection)
            {
                if (errand.PerformerResult == Errand.PerformerResultValue.NeedCorrection)
                {
                    if (System.Windows.Forms.MessageBox.Show(("Вы уже информировали Контролёра о необходимости коррекции, хотите уведомить его повторно?"),
                                                             "Вы нажали кнопку - \"Требуется Коррекция\"", System.Windows.Forms.MessageBoxButtons.OKCancel, System.Windows.Forms.MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.Cancel)
                    { return; }
                }
                //Создаём запрос коррекции
                CorrectionRequest newCorrection = null;

                if (errand.ActualCorrectionRequest == null)
                {
                    newCorrection = new CorrectionRequest(errand);
                }
                else { newCorrection = errand.ActualCorrectionRequest; }

                if (!newCorrection.ShowDialog()) { return; }
                else
                {
                    errand.ActualCorrectionRequest = newCorrection;
                }
                errand.PerformerResult = Errand.PerformerResultValue.NeedCorrection;
                errand.PerformerComment = errand.ActualCorrectionRequest.Reason;
                DateTime date = DateTime.Now;
                //записываем в лог
                try
                {
                    string actionName = string.Format("{0}", performerResult);
                    string message = errand.ActualCorrectionRequest.ToString();
                    if (!AddRecordToLog(errand, date, actionName, Errand.Role.Performer, false, message)) return;
                }
                catch (Errand.Exception_CancelAllActivities e)
                { return; }

                errand.SaveSecondaryParametersToDB();

                string textSubj = "Требуется коррекция. Поручение №" + errand.ErrandNumber.ToString() +
                    ". Документ: " + errand.DocumentForErrand + ". Срок: " + errand.PlanEndDate.ToShortDateString() + ". " + errand.Text;
                string textMsg = "В АСУ \"Динамика\", в модуле \"Контроль поручений\", Исполнитель - " + CurrentUserName
                    + ", запросил коррекцию поручения №" + errand.ErrandNumber.ToString() + ". "
                    + "<br>" + errand.ActualCorrectionRequest.ToString();
                InformCheckerAndDirector(errand, textSubj, textMsg);
            }
            else if (performerResult == Errand.PerformerResultValue.NotDone)
            {
                DateTime date = DateTime.Now;
                //записываем в лог
                try
                {
                    string actionName = string.Format("C поручением ознакомлен.", performerResult.ToLower());

                    if (!AddRecordToLog(errand, date, actionName, Errand.Role.Performer, false)) return;
                }
                catch (Errand.Exception_CancelAllActivities e)
                { return; }

                errand.PerformerResult = Errand.PerformerResultValue.NotDone;
                errand.PerformerAcceptDate = date;

                errand.SaveSecondaryParametersToDB();
            }
            else { throw new ApplicationException("Не приемлемое значение PerformerResult - " + performerResult); }
        }
#endif
        void ReturnToRework(Errand errand)
        {
            if (errand.Stage == Errand.ErrandStages.Closed)
            {
                System.Windows.Forms.MessageBox.Show("Поручение находится в стадии - \"Закрыто\", изменение невозможно.");
                return;
            }

            DateTime date = DateTime.Now;
            //записываем в лог
            try
            {
                if (!AddRecordToLog(errand, date, "Возврат на доработку.", Errand.Role.Checker, true)) return;
            }
            catch (Errand.Exception_CancelAllActivities e)
            { return; }
            errand.VerificationResult = Errand.VerificationResultValue.NotVerified;
            errand.VerificationDate = null;
            errand.PerformerResult = Errand.PerformerResultValue.NotDone;
            errand.PerformerDoneDate = null;

            errand.SaveSecondaryParametersToDB();
            string textSubj = "Возврат на доработку. Поручение №" + errand.ErrandNumber.ToString() +
                ". Документ: " + errand.DocumentForErrand + ". Срок: " + errand.PlanEndDate.ToShortDateString() + ". " + errand.Text;
            string textMsg = "В АСУ \"Динамика\", в модуле \"Контроль поручений\", Контролёр - " + CurrentUserName
                + ", возвращает вам на доработку, поручение №" + errand.ErrandNumber.ToString() + ". " + "<br>Комментарий: " + errand.CheckerComment;
            InformPerformer(errand, textSubj, textMsg);
        }
#if !server
        void ChangeNextCheckDate(Errand errand)
        {
            if (errand.Stage == Errand.ErrandStages.Closed)
            {
                System.Windows.Forms.MessageBox.Show("Поручение находится в стадии - \"Закрыто\", изменение невозможно.");
                return;
            }

            DateTime newNextCheckDate = ChooseDateDialog(errand, "Выберите дату контроля", errand.NextCheckDate);
            if (errand.NextCheckDate == newNextCheckDate) return;

            errand.NextCheckDate = newNextCheckDate;

            AddRecordToLog(errand, DateTime.Now, "Дата следующего контроля изменена - " + errand.NextCheckDate.ToShortDateString(), Errand.Role.Checker, false);

            errand.SaveSecondaryParametersToDB();
        }



        void AddToLogWithDialogShow(Errand errand, string action, Errand.Role role)
        {
            if (errand.Stage == Errand.ErrandStages.Closed)
            {
                System.Windows.Forms.MessageBox.Show("Поручение находится в стадии - \"Закрыто\", изменение невозможно.");
                return;
            }

            DateTime date = DateTime.Now;
            //записываем в лог
            try
            {
                if (!AddRecordToLog(errand, date, action, role, true)) return;
            }
            catch (Errand.Exception_CancelAllActivities e)
            { return; }
            errand.SaveSecondaryParametersToDB();
        }

        DateTime ChooseDateDialog(Errand errand, string dialogName, DateTime initialDate)
        {
            DateTime result = initialDate;
            DateTime dateMin = DateTime.Now.Date;

            UIMacroContext uIMacroContext = new UIMacroContext(new MacroContext(ErrandControlReference.Connection), null, null, null, null);
            var dialog = uIMacroContext.CreateInputDialog();
            dialog.Caption = dialogName;
            dialog.AddStringField("Предупреждение", "Дата контроля не может быть позже даты исполнения!", false, false, "", true, 2);
            DateTime minValue = DateTime.Today;
            dialog.SetElementEnabled("Предупреждение", minValue > initialDate);
            dialog.SetElementVisibility("Предупреждение", false);
            dialog.AddDateField(dialogName, initialDate, true, (int)DateTimePickerFormat.Short, "");
            dialog.SetSize(500, 100);



            dialog.FieldValueChanged += (sender, eventArgs) =>
            {
                if (eventArgs.Name == dialogName)
                {

                    if ((DateTime)eventArgs.NewValue < minValue)
                    {
                        dialog.SetElementVisibility("Предупреждение", true);
                    }
                    else dialog.SetElementVisibility("Предупреждение", false);
                }
            };
            if (dialog.Show(uIMacroContext))
            {

                result = dialog.GetValue(dialogName);
                if (result < minValue) result = ChooseDateDialog(errand, dialogName, result);
            }
            else result = errand.NextCheckDate;

            return result;
        }
#endif

        private void InformCheckerAndDirector(Errand errand, string textSubj, string textMsg)
        {
            //отправляем письмо Контролёру
            List<UserReferenceObject> mailRecipients = new List<UserReferenceObject>() { errand.Checker, errand.Director };
            string email = "";
            SendHTMLMessage(textSubj, textMsg, mailRecipients, email, new List<Errand>() { errand });
        }

        private void InformChecker(Errand errand, string textSubj, string textMsg)
        {
            //отправляем письмо Контролёру
            List<UserReferenceObject> mailRecipients = new List<UserReferenceObject>() { errand.Checker };
            string email = "";
            SendHTMLMessage(textSubj, textMsg, mailRecipients, email, new List<Errand>() { errand });
        }

        private void InformPerformer(Errand errand, string textSubj, string textMsg)
        {
            //отправляем письмо Контролёру и исполнителю
            List<UserReferenceObject> mailRecipients = new List<UserReferenceObject>();
            string email = "";
            if (!errand.IsPerformerExternal) mailRecipients.Add(errand.Performer);
            else
                email = errand.ExternalPerformerEMail;
            SendHTMLMessage(textSubj, textMsg, mailRecipients, email, new List<Errand>() { errand });
        }

        private void InformPerformerAndChecker(Errand errand, string textSubj, string textMsg)
        {
            //отправляем письмо Контролёру и исполнителю
            List<UserReferenceObject> mailRecipients = new List<UserReferenceObject>() { errand.Checker };
            string email = "";
            if (!errand.IsPerformerExternal) mailRecipients.Add(errand.Performer);
            else
                email = errand.ExternalPerformerEMail;
            SendHTMLMessage(textSubj, textMsg, mailRecipients, email, new List<Errand>() { errand });
        }
        private void InformPerformerAndDirector(Errand errand, string textSubj, string textMsg)
        {
            //отправляем письмо Контролёру и исполнителю
            List<UserReferenceObject> mailRecipients = new List<UserReferenceObject>() { errand.Director };
            string email = "";
            if (!errand.IsPerformerExternal) mailRecipients.Add(errand.Performer);
            else
                email = errand.ExternalPerformerEMail;
            SendHTMLMessage(textSubj, textMsg, mailRecipients, email, new List<Errand>() { errand });
        }

        bool AddRecordToLog(Errand errand, DateTime actionTime, string actionName, Errand.Role role, bool showDialog, string msg = "")
        {
            string roleComment = "";
#if !server
            if (role != 0 && showDialog)
            {
                switch (role)
                {
                    case Errand.Role.Director:
                        errand.DirectorComment = GetNewCommentFromUser(errand.DirectorComment);
                        roleComment = errand.DirectorComment;
                        break;
                    case Errand.Role.Performer:
                        errand.PerformerComment = GetNewCommentFromUser(errand.PerformerComment);
                        roleComment = errand.PerformerComment;
                        break;
                    case Errand.Role.Checker:
                        errand.CheckerComment = GetNewCommentFromUser(errand.CheckerComment);
                        roleComment = errand.CheckerComment;
                        break;
                    default: break;
                }
            }
#endif
            errand.Log = "***** " + actionTime.ToString() + "; " + CurrentUserName + "; " + actionName + " *****\r\n"
                + roleComment + "\r\n" + msg + "\r\n" + errand.Log;
            return true;
        }
#if !server
        string GetNewCommentFromUser(string comment)
        {
            //UIMacroContext uIMacroContext = new UIMacroContext(new MacroContext(ErrandControlReference.Connection), null, null, null, null);
            var dialog = this.CreateInputDialog();
            dialog.Caption = "Введите комментарий:";
            dialog.AddString("Комментарий", comment, true, false, false, 5);

            if (dialog.Show())
            {
                comment = dialog.GetValue("Комментарий");
            }
            else { throw new Errand.Exception_CancelAllActivities("Ввод комментария отменён."); }
            return comment;
        }

        /// <summary>
        /// Не работает в новом клиенте
        /// </summary>
        /// <param name="comment"></param>
        /// <returns></returns>
        string GetNewCommentFromUserOld(string comment)
        {
            UIMacroContext uIMacroContext = new UIMacroContext(new MacroContext(ErrandControlReference.Connection), null, null, null, null);
            var dialog = uIMacroContext.CreateInputDialog();
            dialog.Caption = "Введите комментарий:";
            dialog.AddStringField("Комментарий", comment, true, false, "", true, 2);

            if (dialog.Show(uIMacroContext))
            {
                comment = dialog.GetValue("Комментарий");
            }
            else { throw new Errand.Exception_CancelAllActivities("Ввод комментария отменён."); }
            return comment;
        }
#endif
        #endregion
        private List<Errand> GetChangedErrandsFromDate(DateTime last_Send_Date)
        {
            List<ReferenceObject> findedObjects = new List<ReferenceObject>();
            ReferenceInfo referenceInfo = ErrandControlReference.Reference.ParameterGroup.ReferenceInfo;
            //Формирование фильтра
            using (Filter filter = new Filter(referenceInfo))
            {
                Stage stage_В_работе = Stage.Find(Context.Connection, Errand.ErrandStages.New);
                //стадия != Новое поручение
                ReferenceObjectTerm term1 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
                term1.Path.AddParameter(ErrandControlReference.Reference.ParameterGroup.OneToOneParameters.Find(SystemParameterType.Stage));
                term1.Operator = ComparisonOperator.NotEqual;
                term1.Value = stage_В_работе;

                //Дата последнего изменения позднее last_Send_Date
                ReferenceObjectTerm term2 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
                term2.Path.AddParameter(ErrandControlReference.Reference.ParameterGroup[SystemParameterType.EditDate]);
                term2.Operator = ComparisonOperator.GreaterThanOrEqual;
                term2.Value = last_Send_Date;

                //Автор последнего изменения не Киселёва
                ReferenceObjectTerm term4 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
                term4.Path.AddParameter(ErrandControlReference.Reference.ParameterGroup[SystemParameterType.Editor]);
                term4.Operator = ComparisonOperator.NotEqual;
                term4.Value = КиселёваМарина;

                //Руководитель == Островой
                ReferenceObjectTerm term3_1 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
                term3_1.Path.AddGroup(referenceInfo.Description.OneToOneLinks.Find(link_Director_GUID));
                term3_1.Operator = ComparisonOperator.Equal;
                term3_1.Value = Островой;

                //Руководитель == Вахромова
                ReferenceObjectTerm term3_2 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.Or);
                term3_2.Path.AddGroup(referenceInfo.Description.OneToOneLinks.Find(link_Director_GUID));
                term3_2.Operator = ComparisonOperator.Equal;
                term3_2.Value = ВахромоваЯна;

                //Руководитель == Савченкова Мария
                ReferenceObjectTerm term3_3 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.Or);
                term3_3.Path.AddGroup(referenceInfo.Description.OneToOneLinks.Find(link_Director_GUID));
                term3_3.Operator = ComparisonOperator.Equal;
                term3_3.Value = СавченковаМария;


                filter.Terms.GroupTerms(new Term[] { term3_1, term3_2, term3_3});

                //Применяем фильтр
                findedObjects = ErrandControlReference.Reference.Find(filter);

                //System.Windows.Forms.MessageBox.Show(filter.ToString());
            }
            List<Errand> errands = new List<Errand>();
            if (findedObjects != null)
            {
                foreach (var item in findedObjects)
                {
                    errands.Add(new Errand(item));
                }
            }
            return errands;
        }


        const int numDaysForOverdueValidation = 14;
        readonly Guid item_CompanyCalendar_GUID = new Guid("561b03ce-feb0-4f4f-bf20-8c76e2f14bbe");//Guid объекта "Календарь АО ЦНТУ Динамика"


        /// <summary>
        /// возвращает дату numWorkingDays - рабочих дней назад, даже если она попадает на выходной
        /// </summary>
        public DateTime Date_WorkingDaysAgo(DateTime startDate, int numWorkingDays)
        {
            CalendarReferenceObject calendar = new CalendarReference(Connection).Find(item_CompanyCalendar_GUID) as CalendarReferenceObject;
            List<DateTime> workingDaysList = new List<DateTime>();
            int workingDaysCount = 0; //счётчик рабочих дней
            DateTime result = startDate;
            while (workingDaysCount < numWorkingDays)
            {
                bool DayIsWorking = true;

                foreach (var workTime in calendar.WorkingTimes)
                {
                    if (workTime.ContainsDate(result) && workTime.WorkTimeIntervals.Count() == 0) DayIsWorking = false;
                }
                if (DayIsWorking) workingDaysCount++; //если рабочий день, прибавляем к счётчику 1
                                                      //System.Windows.Forms.MessageBox.Show(result.ToString() + "\n" + DayIsWorking.ToString());
                result = result.AddDays(-1).Date;
            }
            return result;
        }




        private bool CorrectionNotNeeded(Errand errand)
        {
            if (errand.Stage == Errand.ErrandStages.InWork
                && (errand.PerformerResult == Errand.PerformerResultValue.NeedCorrection
                    || errand.VerificationResult == Errand.VerificationResultValue.NeedCorrection))
            {
                return false;
            }
            return true;
        }

        private List<CorrectionRequest> GetListOfOverdueCorrectionRequests(int numDaysForOverdueValidation)
        {
            List<ReferenceObject> listOverdueCorrectionRequests = new List<ReferenceObject>();
            //Получаем дату рабочих дней назад(дата отсчёта берётся вчера, т.к. запуск макроса происходит в 00-00)
            DateTime overdueDate = Date_WorkingDaysAgo(DateTime.Today.Date.AddDays(-1), numDaysForOverdueValidation);
            //Формирование фильтра
            using (Filter filter = new Filter(CorrectionRequestsReferenceInfo))
            {
                //дата применения - пустое значение
                ReferenceObjectTerm term11 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
                term11.Path.AddParameter(CorrectionRequestsReferenceInfo.Description.OneToOneParameters.Find(CR_param_ApplyDate_GUID));
                term11.Operator = ComparisonOperator.IsNull;
                //term2.Value = overdueDate;

                //дата применения - пустое значение (minvalue)
                ReferenceObjectTerm term12 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.Or);
                term12.Path.AddParameter(CorrectionRequestsReferenceInfo.Description.OneToOneParameters.Find(CR_param_ApplyDate_GUID));
                term12.Operator = ComparisonOperator.Equal;
                term12.Value = DateTime.MinValue;
#if !test
                //Дата создания раньше срока просрочки
                ReferenceObjectTerm term2 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
                term2.Path.AddParameter(CorrectionRequestsReferenceInfo.Description[SystemParameterType.CreationDate]);
                term2.Operator = ComparisonOperator.LessThan;
                term2.Value = overdueDate;
#endif
                // Группируем условия в отдельную группу (другими словами добавляем скобки)
                TermGroup group1 = filter.Terms.GroupTerms(new Term[] { term11, term12 });

                //filter.Terms.GroupTerms(new TermGroupItem[] { group1, group2 });

                //Применяем фильтр
                listOverdueCorrectionRequests = CorrectionRequestsReference.Find(filter);
                //System.Windows.Forms.MessageBox.Show(filter.ToString());
            }
            //System.Windows.Forms.MessageBox.Show(listOverdueCorrectionRequests.Count.ToString());
            List<CorrectionRequest> result = new List<CorrectionRequest>();
            foreach (var cr_ro in listOverdueCorrectionRequests)
            {
                CorrectionRequest cr = new CorrectionRequest(cr_ro);
                bool hasRequisiteChangeNoticeForPlanDate = cr.RequisiteChangeNotices.Where(rcn => rcn.ChangingParameterName == Errand.Requisites.PlanDate
#if !test
                                                                                           && rcn.IsApproveByChecker
                                                                                           && ((DateTime)(rcn.CheckerApproveDate)) < overdueDate
#endif
                                                                                          ).Count() > 0;
                if (hasRequisiteChangeNoticeForPlanDate)
                    result.Add(cr);
            }
            //System.Windows.Forms.MessageBox.Show("result.Count" + result.Count.ToString());
            return result;
        }

        #region События запускаемые на сервере

        //рассылка списка непринятых поручений, для конкретного Исполнителя
        public void ЕжедневнаяРассылкаОНеПринятыхПоручениях()
        {
            ReferenceInfo referenceInfo = ErrandControlReference.Reference.ParameterGroup.ReferenceInfo;
            Reference reference = ErrandControlReference.Reference;
            //Начало сегодняшнего дня
            DateTime today = DateTime.Now.Date;

            //Stage stage_НовоеПоручение = TFlex.DOCs.Model.Stages.Stage.Find(Context.Connection, "Новый объект");
            //Stage stage_Закрыто = TFlex.DOCs.Model.Stages.Stage.Find(Context.Connection, "Закрыто");
            Stage stage_В_работе = TFlex.DOCs.Model.Stages.Stage.Find(Context.Connection, Errand.ErrandStages.InWork);
            //Stage stage_Коррекция = TFlex.DOCs.Model.Stages.Stage.Find(Context.Connection, "Корректировка");

            List<ReferenceObject> findedObjects = new List<ReferenceObject>();

            //Формирование фильтра
            using (Filter filter = new Filter(referenceInfo))
            {
                /*//Плановая дата выполнения поручения >= заданного
                filter.Terms.AddTerm(reference.ParameterGroup.OneToOneParameters.Find(param_ErrandEndDate_GUID),
                                     ComparisonOperator.GreaterThanOrEqual, today);
                */
                //стадия == В работе_Контроль поручений
                filter.Terms.AddTerm(reference.ParameterGroup.OneToOneParameters.Find(SystemParameterType.Stage),
                                     ComparisonOperator.Equal, stage_В_работе);
                //Результат исполнения поручения == Не принято в работу
                filter.Terms.AddTerm(reference.ParameterGroup.OneToOneParameters.Find(param_PerformerResult_GUID),
                                     ComparisonOperator.Equal, Errand.PerformerResultValue.NotAccepted);
                //результат валидации != Коррекция
                filter.Terms.AddTerm(reference.ParameterGroup.OneToOneParameters.Find(param_ValidationResult_GUID),
                                     ComparisonOperator.NotEqual, Errand.ValidationResultValue.Correction);
                /*//Должность Контролёра == controlerPosition
                filter.Terms.AddTerm(reference.ParameterGroup.OneToOneLinks.Find(link_PositionChecker_GUID).ToString(),
                                     ComparisonOperator.Equal, controlerPosition);*/
                //Применяем фильтр
                findedObjects = reference.Find(filter);
            }
            if (findedObjects == null) return;

            //создаём список исполнителей и поручений
            List<UserReferenceObject> performers = new List<UserReferenceObject>();
            List<Errand> errands = new List<Errand>();
            foreach (ReferenceObject errand_ro in findedObjects)
            {
                Errand errand = new Errand(errand_ro);
                errands.Add(errand);
                if (errand.IsPerformerExternal) continue;
                else
                {
                    if (errand.Performer != null) performers.Add(errand.Performer);
                }
            }
            //делаем выборки поручений по исполнителям и отправляем им письма
            foreach (UserReferenceObject performer in performers.Distinct())
            {
                if (performer == null) continue;
                //MessageBox.Show(performer.ToString());
                List<Errand> performerErrands = new List<Errand>();
                if (performer != null)
                {
                    foreach (var errand in errands)
                    {
                        //проверяем, внешнее ли поручение
                        if (errand.IsPerformerExternal) continue;

                        if (performer == errand.Performer) performerErrands.Add(errand);
                    }
                    if (performerErrands == null || performerErrands.Count == 0) continue;
                    //SendMsgs_Perf_AboutNotAcceptedErrands(performer, errands);
                    string textSubj = "Поручения, с которыми вы ещё не ознакомились.";
                    string textMsg = "В АСУ\"Динамика\", в модуле \"Контроль поручений\", на дату - " + DateTime.Today.Date.ToString("dd.MM.yyyy")
                        + ", есть поручения, с которыми вы ещё не ознакомились, в которых вы отмечены как Исполнитель.";
                    List<ReferenceObject> attach = new List<ReferenceObject> { new FileReference(ErrandControlReference.Connection).Find(file_PerformShortInstr_GUID) };
                    SendHTMLMessage(textSubj, textMsg, new List<UserReferenceObject> { performer }, null, performerErrands, attach);
                    //MessageBox.Show("Send to "+performer.ToString());
                }
            }
        }

        //рассылка списка заканчивающихся сегодня поручений, для конкретного контролёра
        public void ЕжедневнаяРассылкаОЗаканчивающихсяПоручениях()
        {

            ReferenceInfo referenceInfo = ErrandControlReference.Reference.ParameterGroup.ReferenceInfo;
            Reference reference = ErrandControlReference.Reference;
            //Начало сегодняшнего дня
            DateTime today = DateTime.Now.Date;

            Stage stage_В_работе = TFlex.DOCs.Model.Stages.Stage.Find(Context.Connection, Errand.ErrandStages.InWork);

            List<ReferenceObject> findedObjects = new List<ReferenceObject>();

            //Формирование фильтра
            using (Filter filter = new Filter(referenceInfo))
            {
                //Плановая дата выполнения поручения >= заданного
                filter.Terms.AddTerm(reference.ParameterGroup.OneToOneParameters.Find(param_ErrandEndDate_GUID),
                                     ComparisonOperator.GreaterThanOrEqual, today);
                //Плановая дата выполнения поручения < завтра
                filter.Terms.AddTerm(reference.ParameterGroup.OneToOneParameters.Find(param_ErrandEndDate_GUID),
                                     ComparisonOperator.LessThan, today.AddDays(1));
                //стадия == В работе_Контроль поручений
                filter.Terms.AddTerm(reference.ParameterGroup.OneToOneParameters.Find(SystemParameterType.Stage),
                                     ComparisonOperator.Equal, stage_В_работе);
                //результат валидации != Коррекция
                filter.Terms.AddTerm(reference.ParameterGroup.OneToOneParameters.Find(param_ValidationResult_GUID),
                                     ComparisonOperator.NotEqual, "Коррекция");
                //Применяем фильтр
                findedObjects = reference.Find(filter);
                //MessageBox.Show(filter.ToString());
            }
            if (findedObjects == null) return;
            //создаём список контролёров
            List<UserReferenceObject> checkers = new List<UserReferenceObject>();
            List<Errand> errands = new List<Errand>();
            foreach (ReferenceObject item in findedObjects)
            {
                Errand errand = new Errand(item);
                errands.Add(errand);
                if (errand.Checker != null)
                {
                    checkers.Add(errand.Checker);
                }
            }

            foreach (UserReferenceObject checker in checkers.Distinct())
            {
                List<Errand> checkerErrands = new List<Errand>();
                if (checker != null)
                {
                    foreach (var errand in errands)
                    {
                        //добавляем в список поручений для конкретного контролёра не проверенные и не выполненные поручения
                        if (checker == errand.Checker && errand.VerificationResult != Errand.VerificationResultValue.Match) checkerErrands.Add(errand);//
                    }
                    if (checkerErrands == null || checkerErrands.Count == 0) continue;
                    //SendMsgsAboutEndPlanDate(controler, errands);
                    string textSubj = "Сегодня заканчивается срок выполнения у следующих поручений.";
                    string textMsg = "В АСУ\"Динамика\", в модуле \"Контроль поручений\", сегодня " + DateTime.Today.Date.ToString("dd.MM.yyyy")
                        + ", заканчивается плановый срок у следующих поручений в которых вы назначены Контролёром:";
                    //List<ReferenceObject> attach = new List<ReferenceObject>{new FileReference(servConnect).Find(file_PerformShortInstr_GUID)};
                    SendHTMLMessage(textSubj, textMsg, new List<UserReferenceObject> { checker }, null, checkerErrands);
                }
            }

            //создаём список исполнителей
            List<UserReferenceObject> performers = new List<UserReferenceObject>();
            foreach (Errand errand in errands)
            {
                if (errand.Performer != null) performers.Add(errand.Performer);
            }
            //MessageBox.Show(findedObjects.Count.ToString());
            foreach (UserReferenceObject performer in performers.Distinct())
            {
                //MessageBox.Show(performer.ToString());
                List<Errand> performerErrands = new List<Errand>();
                if (performer != null)
                {
                    foreach (var errand in errands)
                    {
                        //проверяем, внешнее ли поручение
                        if (errand.IsPerformerExternal) continue;
                        //добавляем в список поручений для конкретного Исполнителя все за исключением выполненных поручений
                        if (performer == errand.Performer && errand.PerformerResult != Errand.PerformerResultValue.Done) performerErrands.Add(errand);
                    }
                    if (performerErrands == null || performerErrands.Count == 0) continue;
                    //SendMsgs_Perf_AboutEndPlanDate(performer, errands);
                    string textSubj = "Сегодня заканчивается срок выполнения у следующих поручений.";
                    string textMsg = "В АСУ\"Динамика\", в модуле \"Контроль поручений\", сегодня " + DateTime.Today.Date.ToString("dd.MM.yyyy")
                        + ", заканчивается плановый срок у следующих поручений в которых вы назначены Исполнителем:";
                    //List<ReferenceObject> attach = new List<ReferenceObject>{new FileReference(servConnect).Find(file_PerformShortInstr_GUID)};
                    SendHTMLMessage(textSubj, textMsg, new List<UserReferenceObject> { performer }, null, performerErrands);
                }
            }

            //создаём список Руководителей
            List<UserReferenceObject> directors = new List<UserReferenceObject>();
            foreach (Errand errand in errands)
            {
                if (errand.Director != null) directors.Add(errand.Director);
            }

            foreach (UserReferenceObject director in directors.Distinct())
            {
                List<Errand> directorsErrands = new List<Errand>();
                if (director != null)
                {
                    foreach (var errand in errands)
                    {
                        //добавляем в список поручений для конкретного Руководителя все поручения
                        if (director == errand.Director) directorsErrands.Add(errand);
                    }
                    if (directorsErrands == null || directorsErrands.Count == 0) continue;
                    //MessageBox.Show(errands.Count.ToString());
                    //SendHTMLMessageAboutEndDate(director, errands);
                    string textSubj = "Сегодня заканчивается срок выполнения у следующих поручений.";
                    string textMsg = "В АСУ\"Динамика\", в модуле \"Контроль поручений\", сегодня " + DateTime.Today.Date.ToString("dd.MM.yyyy")
                        + ", заканчивается плановый срок у следующих поручений в которых вы назначены Руководителем:";
                    //List<ReferenceObject> attach = new List<ReferenceObject>{new FileReference(servConnect).Find(file_PerformShortInstr_GUID)};
                    SendHTMLMessage(textSubj, textMsg, new List<UserReferenceObject> { director }, null, directorsErrands);
                }
            }
        }

        public void РассылкаДляКиселёвойОПоследнихИзменениях()
        {
            Guid last_Send_Date_Guid = new Guid("0248aba9-e478-4d16-8969-9ac782f64794");
            GlobalParameter last_Send_Date_GlobalParameter = Connection.References.GlobalParameters.Find(last_Send_Date_Guid) as GlobalParameter;
            DateTime last_Send_Date = (DateTime)(last_Send_Date_GlobalParameter.Value);
#if test
			MessageBox.Show(last_Send_Date.ToString());
#endif
            DateTime new_LastSendDate = DateTime.Now;
            List<Errand> errands = GetChangedErrandsFromDate(last_Send_Date);
#if !test
            List<UserReferenceObject> mailRecipients = new List<UserReferenceObject>();
            if (((User)CurrentUser).IsSystem) mailRecipients.Add((UserReferenceObject)КиселёваМарина);
            else
                mailRecipients.Add((UserReferenceObject)CurrentUser);

            if (errands.Count > 0)
                SendHTMLMessage("Последние изменения", string.Format("Поручения изменённые после {0}", null, last_Send_Date), mailRecipients, null, errands);

            if (((User)CurrentUser).IsSystem)
            {
                last_Send_Date_GlobalParameter.BeginChanges();
                last_Send_Date_GlobalParameter.Value.Value = new_LastSendDate;
                last_Send_Date_GlobalParameter.EndChanges();
            }
#endif
        }

        public void АвтоматическоеПрименениеПереносаСрокаПоКоррекции()
        {
            List<CorrectionRequest> списокПросроченныхЗапросовКоррекции = GetListOfOverdueCorrectionRequests(numDaysForOverdueValidation);

            StringBuilder str = new StringBuilder();
            Dictionary<ReferenceObject, List<Errand>> mailRecipientsAndErrands = new Dictionary<ReferenceObject, List<Errand>>();
            List<Errand> errorErrands = new List<Errand>();
            foreach (var correctionRequest in списокПросроченныхЗапросовКоррекции)
            {

                Errand errand = new Errand(ErrandControlReference.Reference.Find(new Guid(correctionRequest.ErrandGuid)));
                if (CorrectionNotNeeded(errand))
                {
                    correctionRequest.Apply("Коррекция уже не требуется.");
                    continue;
                }
                //str.AppendLine(item[param_ErrandNum_GUID].GetInt16().ToString() + " | " + item[param_DoneTaskDate_GUID].GetDateTime().ToShortDateString() + "|" + item[param_ErrandEndDate_GUID].GetDateTime().ToShortDateString() + "|" + item[param_ValidationResult_GUID].GetString());

                try
                {
                    AutomaticChangePlanDateByCorrectionRequest(errand);

                    AddMailRecipientsToDictionary(errand, mailRecipientsAndErrands);

                }
                catch (Exception e)
                {
                    str.AppendLine(errand.ErrandNumber.ToString());
                    str.AppendLine(e.Message.ToString());
                    errorErrands.Add(errand);
                }
            }
            SendMessagesToRecipients(mailRecipientsAndErrands, "Автоматически скорректированные поручения", "Следующие поручения в соответствии с регламентом были скорректированы по сроку:");
            if (str.Length > 0)
                SendMessageToLipaev(str, errorErrands);
            /*System.Windows.Forms.MessageBox.Show(str.ToString());
            System.Windows.Forms.Clipboard.Clear();
            System.Windows.Forms.Clipboard.SetText(str.ToString());
             */
        }

        public void АвтоматическаяВалидацияПоСроку()
        {
            List<ReferenceObject> списокПоручений = GetListOfErrandsWithOverdueValidation(numDaysForOverdueValidation);
            StringBuilder str = new StringBuilder();
            Dictionary<ReferenceObject, List<Errand>> mailRecipientsAndErrands = new Dictionary<ReferenceObject, List<Errand>>();
            List<Errand> errorErrands = new List<Errand>();
            foreach (var item in списокПоручений)
            {
                Errand errand = new Errand(item);
                //str.AppendLine(item[param_ErrandNum_GUID].GetInt16().ToString() + " | " + item[param_DoneTaskDate_GUID].GetDateTime().ToShortDateString() + "|" + item[param_ErrandEndDate_GUID].GetDateTime().ToShortDateString() + "|" + item[param_ValidationResult_GUID].GetString());
                ReferenceObject actualItem = null;
                try
                {
                    if (item.TryBeginChanges(out actualItem))
                    {
                        DateTime actionTime = DateTime.Now;
                        AddRecordToLog(errand, actionTime, "Автоматическая валидация", (Errand.Role)0, false, "Поручение валидировано по истечении срока давности");
                        errand.ValidationResult = Errand.ValidationResultValue.Done;
                        errand.ValidationDate = actionTime;
                        errand.SaveSecondaryParametersToDB();
                        errand.ChangeStageTo(Errand.ErrandStages.Closed);
                        AddMailRecipientsToDictionary(errand, mailRecipientsAndErrands);
                    }
                }
                catch (Exception e)
                {
                    str.AppendLine(item.SystemFields.Id.ToString());
                    str.AppendLine(e.ToString());
                    errorErrands.Add(errand);
                }
            }
            SendMessagesToRecipients(mailRecipientsAndErrands, "Автоматически валидированные поручения", "Следующие поручения в соответствии с регламентом были отмечены как валидированные:");
            if (str.Length > 0)
                SendMessageToLipaev(str, errorErrands);
            /*System.Windows.Forms.MessageBox.Show(str.ToString());
            System.Windows.Forms.Clipboard.Clear();
            System.Windows.Forms.Clipboard.SetText(str.ToString());
             */
        }

        /// <summary>
        /// не изменять сигнатуру и не удалять, используется для задач на сервере при вызове errand.ChangeStageTo(...)
        /// </summary>
        /// <param name="current_Guid"></param>
        public void ChangeStageClosed(string current_Guid)
        {
            Guid objectGuid = new Guid(current_Guid);
            ReferenceObject current = ErrandControlReference.Reference.Find(objectGuid);
            Stage stage = TFlex.DOCs.Model.Stages.Stage.Find(Context.Connection, Errand.ErrandStages.Closed);
            stage.Change(new List<ReferenceObject> { current });
        }

        /// <summary>
        /// не изменять сигнатуру и не удалять, используется для задач на сервере при вызове errand.ChangeStageTo(...)
        /// </summary>
        /// <param name="current_Guid"></param>
        public void ChangeStageInWork(string current_Guid)
        {
            Guid objectGuid = new Guid(current_Guid);
            ReferenceObject current = ErrandControlReference.Reference.Find(objectGuid);
            Stage stage = TFlex.DOCs.Model.Stages.Stage.Find(Context.Connection, Errand.ErrandStages.InWork);
            stage.Change(new List<ReferenceObject> { current });
        }

        #endregion
        private void SendMessageToLipaev(StringBuilder str, List<Errand> errorErrands)
        {
            UserReferenceObject lipaev = new UserReference(Connection).FindUser("Липаев Алексей Александрович");
            SendHTMLMessage("Поручения с ошибками", "При автоматической валидации в поручениях произошли следующие ошибки:" + str.ToString(), new List<UserReferenceObject>() { lipaev }, null, errorErrands);
        }

        private void SendMessagesToRecipients(Dictionary<ReferenceObject, List<Errand>> mailRecipientsAndErrands, string subject, string message)
        {
            if (mailRecipientsAndErrands == null || mailRecipientsAndErrands.Count <= 0) return;
            foreach (var kvp in mailRecipientsAndErrands)
            {
                List<ReferenceObject> errands_ro = new List<ReferenceObject>();
                List<Errand> errands = kvp.Value.Distinct().ToList();
                SendHTMLMessage(subject, message, new List<UserReferenceObject>() { kvp.Key as UserReferenceObject }, null, errands);
            }
        }



        private void AddMailRecipientsToDictionary(Errand errand, Dictionary<ReferenceObject, List<Errand>> mailRecipientsAndErrands)
        {
            if (!errand.IsPerformerExternal)
            {
                AddMailRecipientToDictionary(errand.Performer, errand, mailRecipientsAndErrands);
            }
            AddMailRecipientToDictionary(errand.Checker, errand, mailRecipientsAndErrands);
            AddMailRecipientToDictionary(errand.Director, errand, mailRecipientsAndErrands);
        }

        private void AddMailRecipientToDictionary(ReferenceObject mailRecipient, Errand errand, Dictionary<ReferenceObject, List<Errand>> mailRecipientsAndErrands)
        {
            List<Errand> errands = new List<Errand>();
            if (mailRecipientsAndErrands.TryGetValue(mailRecipient, out errands))
            {
                errands.Add(errand);
                mailRecipientsAndErrands[mailRecipient] = errands;
            }
            else
            {
                errands = new List<Errand>() { errand };
                mailRecipientsAndErrands.Add(mailRecipient, errands);
            }
        }


        private List<ReferenceObject> GetListOfErrandsWithOverdueValidation(int overdueValidationDays)
        {
            List<ReferenceObject> findedObjects = new List<ReferenceObject>();
            //Получаем дату рабочих дней назад(дата отсчёта берётся вчера, т.к. запуск макроса происходит в 00-00)
            DateTime overdueDate = Date_WorkingDaysAgo(DateTime.Today.Date.AddDays(-1), overdueValidationDays);
            //Формирование фильтра
            ReferenceInfo ErrandsControlReferenceInfo = ErrandControlReference.Reference.ParameterGroup.ReferenceInfo;
            Reference ErrandsControlReference = ErrandControlReference.Reference;
            using (Filter filter = new Filter(ErrandsControlReferenceInfo))
            {
                Stage stage_В_работе = Stage.Find(Context.Connection, Errand.ErrandStages.InWork);
                //стадия == В работе_Контроль поручений
                ReferenceObjectTerm term1 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
                term1.Path.AddParameter(ErrandsControlReference.ParameterGroup.OneToOneParameters.Find(SystemParameterType.Stage));
                term1.Operator = ComparisonOperator.Equal;
                term1.Value = stage_В_работе;

                //Плановая дата верификации поручения >= заданного
                ReferenceObjectTerm term2 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
                term2.Path.AddParameter(ErrandsControlReference.ParameterGroup.OneToOneParameters.Find(param_VerificationDate_GUID));
                term2.Operator = ComparisonOperator.LessThan;
                term2.Value = overdueDate;

                //результат валидации == Не валидировано
                ReferenceObjectTerm term3 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
                term3.Path.AddParameter(ErrandsControlReference.ParameterGroup.OneToOneParameters.Find(param_ValidationResult_GUID));
                term3.Operator = ComparisonOperator.Equal;
                term3.Value = Errand.ValidationResultValue.NotValidated;

                //результат верификации не входит в список "Требуется коррекция", "Не верифицировано"
                ReferenceObjectTerm term41 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
                term41.Path.AddParameter(ErrandsControlReference.ParameterGroup.OneToOneParameters.Find(param_VerificationResult_GUID));
                term41.Operator = ComparisonOperator.IsNotOneOf;
                term41.Value = new List<string>() { Errand.VerificationResultValue.NeedCorrection, Errand.VerificationResultValue.NotVerified };

                //результат исполнения == Выполнено
                ReferenceObjectTerm term42 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
                term42.Path.AddParameter(ErrandsControlReference.ParameterGroup.OneToOneParameters.Find(param_PerformerResult_GUID));
                term42.Operator = ComparisonOperator.Equal;
                term42.Value = Errand.PerformerResultValue.Done;

                //результат верификации равен "Соответствует"
                ReferenceObjectTerm term51 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.Or);
                term51.Path.AddParameter(ErrandsControlReference.ParameterGroup.OneToOneParameters.Find(param_VerificationResult_GUID));
                term51.Operator = ComparisonOperator.Equal;
                term51.Value = Errand.VerificationResultValue.Match;

                //результат исполнения == Выполнено
                ReferenceObjectTerm term52 = new ReferenceObjectTerm(filter.Terms, LogicalOperator.And);
                term52.Path.AddParameter(ErrandsControlReference.ParameterGroup.OneToOneParameters.Find(param_ExtErrand_GUID));
                term52.Operator = ComparisonOperator.Equal;
                term52.Value = true;

                // Группируем условия в отдельную группу (другими словами добавляем скобки)
                TermGroup group1 = filter.Terms.GroupTerms(new Term[] { term41, term42 });
                TermGroup group2 = filter.Terms.GroupTerms(new Term[] { term51, term52 });

                filter.Terms.GroupTerms(new TermGroupItem[] { group1, group2 });


                //Применяем фильтр
                findedObjects = ErrandsControlReference.Find(filter);
                //System.Windows.Forms.MessageBox.Show(filter.ToString());
            }
            return findedObjects;
        }



        void SendHTMLMessage(string textSubj, string textMsg, List<UserReferenceObject> mailRecipients, string eMail, List<Errand> errands, List<ReferenceObject> attachments = null)
        {
            //исключаем из рассылки Острового
            UserReference userRef = new UserReference(ErrandControlReference.Connection);
            UserReferenceObject ostrovoy = userRef.Find(new Guid("e86c95d6-dadd-4839-b5ce-efea2e88b1a3")) as UserReferenceObject;

            string HTML_ErrandsTable = CreateHTMLTableForErrands(errands);
            StringBuilder strBld = new StringBuilder();
            strBld.Append(String.Format(@"
    <p>Здравствуйте!</p>
    <p> {0} </p>
<p> {1} </p>", textMsg, HTML_ErrandsTable));

            //Добавление адресатов
            List<ReferenceObject> tempUsers = new List<ReferenceObject>();
            foreach (UserReferenceObject mailgroup in mailRecipients)
            {
                if (mailgroup.ToString() == CurrentUserName) continue;
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
            if (!string.IsNullOrWhiteSpace(eMail))
            {
                externalEmails.Add(eMail);
            }
            string textMessage = strBld.ToString();
            List<ReferenceObject> attacments = new List<ReferenceObject>();
            foreach (var errand in errands)
            {
                attacments.Add(errand.ReferenceObject);
            }
            HTML_MailMessage message = new HTML_MailMessage(textSubj, strBld.ToString(), users, externalEmails, attacments);
            message.Send();
        }

        string CreateHTMLTableForErrands(List<Errand> errands)
        {
            StringBuilder strBld = new StringBuilder();

            strBld.Append(@"<!--Container Table-->
    <table  cellpadding = ""0"" cellspacing = ""0"" border = ""0"" width = ""99 % "">
             <tr>
               <!--Container Table-->
           <table  cellpadding = ""5"" cellspacing = ""0"" border = ""1"" width = ""10 % "">
                    <tr><td colspan = ""4"" bgcolor =#e1e1e1 align=""center"">Расцветка для комментариев Контролёра и Исполнителя</td></tr>
  <tr>
                         <td bgcolor =#62ebff style=align=""center""  nowrap> &nbsp;Выполнено&nbsp; </td>
  <td bgcolor =#ff859c style=align=""center""  nowrap> &nbsp;Не выполнено/Не принято к исполнению&nbsp; </td>
  <td style = align = ""center"" nowrap> &nbsp; Не проверено/ В работе </td>
          <td bgcolor =#fff962 style=align=""center"" nowrap> &nbsp;Требуется коррекция&nbsp; </td>
  </tr>
           </table>
           <br>
         
           <!--End Container Table -->
             </tr>
                 <tr>
                     <td align = ""left""></td>
                        <!--End Email Wrapper Table-->
                        <table  cellpadding = ""0"" cellspacing = ""0"" border = ""1"" width = ""99 % "">
                     
                                  <tr>
                                     <th nowrap bgcolor =#e1e1e1> № </th>
				<th nowrap bgcolor =#e1e1e1> Текст поручения </th>
				<th nowrap bgcolor =#e1e1e1> Срок выполнения </th>
				<th nowrap bgcolor =#e1e1e1> Исполнитель </th>
				<th nowrap bgcolor =#e1e1e1> Контролёр </th>
				<th nowrap bgcolor =#e1e1e1> Руководитель </th>
				<th nowrap bgcolor =#e1e1e1> Ожидаемый результат </th>
				<th bgcolor =#e1e1e1> Комментарий Контролёра </th>
				<th bgcolor =#e1e1e1> Комментарий Исполнителя </th>
			 </tr> ");

            foreach (var errand in errands)
            {
                string bgColorContr = "#0ffda1";
                switch (errand.VerificationResult)
                {
                    case Errand.VerificationResultValue.Match:
                        bgColorContr = "#62ebff";
                        break;
                    case Errand.VerificationResultValue.DoesNotMatch:
                        bgColorContr = "#ff859c";
                        break;
                    case Errand.VerificationResultValue.NotVerified:
                        bgColorContr = "#ffffff";
                        break;
                    case Errand.VerificationResultValue.NeedCorrection:
                        bgColorContr = "#fff962";
                        break;
                    default: break;
                }

                //по умолчанию результат - не принято в работу
                string bgColorPerf = "#ff859c";
                switch (errand.PerformerResult)
                {
                    case Errand.PerformerResultValue.Done:
                        bgColorPerf = "#62ebff";
                        break;
                    case Errand.PerformerResultValue.NotDone:
                        bgColorPerf = "#ff859c";
                        break;
                    case Errand.PerformerResultValue.NeedCorrection:
                        bgColorPerf = "#fff962";
                        break;
                    default: break;
                }
                string performer;
                if (errand.IsPerformerExternal) performer = errand.ExternalPerformer;
                else performer = errand.Performer.ToString();
                string commentPerform = errand.PerformerComment;
                if (errand.PerformerDoneDate != null)
                    commentPerform =
                        ((DateTime)errand.PerformerDoneDate).ToString("dd.MM.yyyy") +
                        "<br>" +
                        errand.PerformerComment;
                string commentControl = errand.CheckerComment;
                if (errand.VerificationDate != null)
                    commentControl =
                        ((DateTime)errand.VerificationDate).ToString("dd.MM.yyyy") +
                        "<br>" +
                        errand.CheckerComment;
                string link = HTML_MailMessage.GetLinkFor(errand.ReferenceObject);
                strBld.Append(String.Format(@"
<tr>
<td>{10}</td>
<td style=align=""right""><a href={0}>{1}</a></td>
<td>{11}</td>
<td style=align=""left"">{2}</td>
<td>{8}</td>
<td>{9}</td>
<td>{3}</td>
<td bgcolor=""{4}"">{5}</td>
<td bgcolor=""{6}"">{7}</td></tr>",
                                            link,
                                            errand.Text + "&nbsp",
                                            performer + "&nbsp",
                                            errand.ProposedResult + "&nbsp",
                                            bgColorContr,
                                            commentControl + "&nbsp",
                                            bgColorPerf,
                                            commentPerform + "&nbsp",
                                            errand.Checker.ToString() + "&nbsp",
                                            errand.Director.ToString() + "&nbsp",
                                            errand.ErrandNumber.ToString(),
                                            errand.PlanEndDate.ToString("dd.MM.yyyy")));
            }
            strBld.Append(@"</table>");

            return strBld.ToString();
        }






#if !server
        public void ShowCoef()
        {
            ReferenceObject errand_ro = Context.ReferenceObject;
            Errand newErrand = new Errand(errand_ro);
            MessageBox.Show(newErrand.CoefficientOfExecution.ToString());
        }

        public void CreateNewErrand(string errandText, string documentForErrand, string proposedResult, DateTime planEndDate, User director, User performer, User checker, IEnumerable<ReferenceObject> initialData = null)
        {
            Errand newErrand = new Errand();
            //Guid ON_param_Content_Guid = new Guid("62d753ad-1c1e-4c67-b4d4-f631d70864d8");//гуид параметра "содержание служебной записки" справочника "Служебные записки"
            newErrand.Text = errandText;
            newErrand.DocumentForErrand = documentForErrand;
            newErrand.ProposedResult = proposedResult;
            //MessageBox.Show((string)ВыполнитьМакрос("Служебные записки", "GetRegistryNumberAndDateON", officialNote));
            foreach (var item in initialData)
            {
                newErrand.AddInitialData(item);
            }

            newErrand.SaveRequisitesToDB();
            RefObj newRefObj = RefObj.CreateInstance(newErrand.ReferenceObject, Context);

            //var uiContext = Context as UIMacroContext;
            //uiContext.CloseDialog(true); // Сохраняем изменения и закрываем диалог

            ShowPropertyDialog(newRefObj);
        }

        public void СоздатьПоручениеПоСлужебнойЗаписке()
        {
            ReferenceObject officialNote = Context.ReferenceObject;
            Errand newErrand = new Errand();
            //Guid ON_param_Content_Guid = new Guid("62d753ad-1c1e-4c67-b4d4-f631d70864d8");//гуид параметра "содержание служебной записки" справочника "Служебные записки"
            newErrand.Text = (string)ВыполнитьМакрос("Служебные записки", "GetTextForApprovedON", officialNote);
            newErrand.DocumentForErrand = (string)ВыполнитьМакрос("Служебные записки", "GetRegistryNumberAndDateON", officialNote);
            //MessageBox.Show((string)ВыполнитьМакрос("Служебные записки", "GetRegistryNumberAndDateON", officialNote));
            newErrand.AddInitialData(officialNote);
            ReferenceObject reportFile = (ReferenceObject)ВыполнитьМакрос("Служебные записки", "GetReportFileForApprovedON", officialNote);
            if (reportFile != null) newErrand.AddInitialData(reportFile);
            foreach (var file in ВыполнитьМакрос("Служебные записки", "GetAdditionalFilesForApprovedON", officialNote))
            {
                newErrand.AddInitialData((ReferenceObject)file);
            }
            newErrand.SaveSecondaryParametersToDB();
            RefObj newRefObj = RefObj.CreateInstance(newErrand.ReferenceObject, Context);

            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true); // Сохраняем изменения и закрываем диалог

            ShowPropertyDialog(newRefObj);
        }

        /// <summary>
        /// Кнопка "Перепоручить" - создаёт копию текущего поручения и открывает его свойства
        /// </summary>
        public bool Перепоручить()
        {
            ReferenceObject current = Context.ReferenceObject;
            ReferenceObject newErrand = current.CreateCopy();
            ForCopyClearParams(newErrand);
            newErrand.SetLinkedObject(link_Director_GUID, (ReferenceObject)CurrentUser);
            newErrand.SetLinkedObject(link_Checker_GUID, (ReferenceObject)CurrentUser);
            newErrand.SetLinkedObject(link_Performer_GUID, null);
            RefObj newRefObj = RefObj.CreateInstance(newErrand, Context);

            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true); // Сохраняем изменения и закрываем диалог

            if (ShowPropertyDialog(newRefObj)) return true;
            return false;
        }


        /* К удалению, по моему нигде не используется 13/12/2017
        public List<User> Directors(ReferenceObject errand)
        {
            List<User> result = new List<User>();
            User person = errand.GetObject(link_Director_GUID) as User;
            if (person == null) return null;
            result.Add(person);
            result.AddRange(person.GetAllInternalUsers());
            return result;
        }

        public List<User> Checkers(ReferenceObject errand)
        {
            List<User> result = new List<User>();
            User person = errand.GetObject(link_Checker_GUID) as User;
            if (person == null) return null;
            result.Add(person);
            result.AddRange(person.GetAllInternalUsers());
            return result;
        }

        public List<User> Performers(ReferenceObject errand)
        {
            List<User> result = new List<User>();
            User person = errand.GetObject(link_Performer_GUID) as User;
            if (person == null) return null;
            result.Add(person);
            result.AddRange(person.GetAllInternalUsers());
            return result;
        }

        */
        public void AddNumber()
        {
            ГлобальныйПараметр["Номер_поручения"] += 1;
            Параметр["Номер поручения"] = ГлобальныйПараметр["Номер_поручения"];
            Параметр["Плановая дата выполнения поручения"] = DateTime.Now.AddDays(30);
            ReferenceObject current = Context.ReferenceObject;
            if (current.GetObject(link_Director_GUID) == null) current.SetLinkedObject(link_Director_GUID, (ReferenceObject)CurrentUser);
            if (current.GetObject(link_Checker_GUID) == null) current.SetLinkedObject(link_Checker_GUID, (ReferenceObject)CurrentUser);
            current.ApplyChanges();
        }

        #region События

        #endregion


        void ForCopyClearParams(ReferenceObject current)
        {
            Errand errand = new ErrandControl.Errand(current);
            if (!string.IsNullOrWhiteSpace(errand.Log) ||
                errand.PerformerResult != Errand.PerformerResultValue.NotAccepted ||
                errand.VerificationResult != Errand.VerificationResultValue.NotVerified ||
                errand.ValidationResult.ToString() != Errand.ValidationResultValue.NotValidated)
            {

                //стираем Журнал событий
                errand.Log = "";
                //проверяем внешнее поручение или нет
                //меняем результаты на начальные
                errand.PerformerResult = Errand.PerformerResultValue.NotAccepted;
                errand.VerificationResult = Errand.VerificationResultValue.NotVerified;
                errand.ValidationResult = Errand.ValidationResultValue.NotValidated;
                //Стираем комментарии
                errand.DirectorComment = "";
                errand.CheckerComment = "";
                errand.PerformerComment = "";
                //обнуляем даты
                errand.SendToPerformerDate = null;
                errand.PerformerAcceptDate = null;
                errand.PerformerDoneDate = null;
                errand.VerificationDate = null;
                errand.ValidationDate = null;
                errand.NextCheckDate = errand.PlanEndDate;
                errand.SaveSecondaryParametersToDB();
            }
        }

        public void NewErrandFromOffice2()
        {
            ServerConnection servConnect = (Context as MacroContext).Connection;

            // создание объекта для работы с данными
            Reference referenceUsers = UserReference;

            //var должность;
            ReferenceObject position = null;
            Объект подразделение = null;
            ReferenceObject department = null;

            Объект исполнитель = ТекущийОбъект.СвязанныйОбъект["Ответственный - Группы и пользователи"];
            if (исполнитель != null)
            {
                var должность = ВыполнитьМакрос("6a4385ff-475a-4afb-bb85-965252abebdc", "ОбъектДолжность", исполнитель);
                //Guid должностьGuid = new Guid(должность.Параметр["Guid"].ToString());
                //position = referenceUsers.Find(должностьGuid);
                подразделение = ВыполнитьМакрос("6a4385ff-475a-4afb-bb85-965252abebdc", "Подразделение", должность);
                //Guid подразделениеGuid = new Guid(подразделение.Параметр["Guid"].ToString());
                department = (ReferenceObject)подразделение;
            }

            ReferenceObject rkk = Context.ReferenceObject;
            //рассылка для ознакомления
            // получение списка пользователей для ознакомления
            List<ReferenceObject> users = rkk.GetObjects(guid_komu);
            string nomer = "";
            nomer = rkk.GetObjectValue("Регистрационный номер").ToString();
            string org = "";
            org = rkk.GetObjectValue("Откуда поступил").ToString();

            List<User> mailRecipients = new List<User>();

            foreach (ReferenceObject user in users)
            {
                //Добавляем адресатов сообщения
                User mailuser = user as User;
                if (mailuser != null)
                    mailRecipients.Add(mailuser);
            }
            mailRecipients.Distinct();

            // поиск типа
            ClassObject classObject = ErrandControlReference.Errand_Class;
            // создание
            ReferenceObject newObject = ErrandControlReference.Reference.CreateReferenceObject(classObject);
            // установка связи 1:1 или N:1 ответственный - исполнитель
            newObject.SetLinkedObject(link_Performer_GUID, rkk.GetObject(link_ToResponsible_GUID));
            // установка связи 1:1 или N:1 Должность Исполнителя
            newObject.SetLinkedObject(link_PositionPerformer_GUID, position);
            // установка связи 1:1 или N:1 Подразделение Исполнителя
            if (department != null)
                newObject.SetLinkedObject(link_DepartmentPerformer_GUID, department);
            //ЗаполнитьИсполнителя();
            newObject[param_ErrandDocNum_GUID].Value = "Входящая корреспонденция №" + rkk[param_RegNum_GUID].Value.ToString();
            newObject.AddLinkedObject(link_ToInputData_GUID, rkk);
            newObject[param_Category_GUID].Value = "Входящая корреспонденция";
            foreach (ReferenceObject item in rkk.GetObjects(link_ToDocuments_GUID))
            {
                newObject.AddLinkedObject(link_ToInputData_GUID, item);
            }
            newObject.ApplyChanges();

            RefObj newRefObj = RefObj.CreateInstance(newObject, Context);

            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true); // Сохраняем изменения и закрываем диалог

            if (ShowPropertyDialog(newRefObj))
            {

                MailMessage message = new MailMessage(servConnect.Mail.DOCsAccount); //Создаем новое сообщение
                message.Subject = "Вам перенаправлена входящая корреспонденция N " + nomer + " от " + org;
                message.Body = "1. Одинарным кликом откройте вложенный объект, который находится справа от иконки \"Сообщение\" \r\n"
                    + "2. В открывшемся диалоге перейдите на вкладку \"Журнал ознакомления\" \r\n"
                    + "3. Нажмите на кнопку \"Оставить запись об ознакомлении\" \r\n4. В открывшемся диалоге подтвердите ознакомление нажав кнопку ОК\r\n\r\n"
                    + "Ответственный:   " + rkk.GetObject(link_ToResponsible_GUID).ToString();
                foreach (User user in mailRecipients)
                {
                    //Добавляем адресатов сообщения
                    //User mailuser = user as User;
                    message.To.Add(new MailUser(user));
                }
                //Прикрепляем к сообщению объект справочника
                message.Attachments.Add(new ObjectAttachment(rkk));
                message.Send(); //Отправляем сообщение

            }
        }
        #region События

        #region События справочников


        /// <summary>
        /// Блокирует изменение параметров на стадии В работе (Прозрачное редактирование)
        /// </summary>
        public bool DisableChangingRequisites()
        {
            //ProjectManagement.Refresh();
            //ServerConnection servConnect = (Context as MacroContext).Connection;
            ReferenceObject current = Context.ReferenceObject;
            var changedParam = Context.ChangedParameter;
            if (changedParam == null) return false;
            Stage stage = TFlex.DOCs.Model.Stages.Stage.Find(Connection, "В работе_Контроль поручений");
            if (current.SystemFields.Stage != stage) return false;
            //если поручение корректируется, то редактирование доступно
            if (current[param_ValidationResult_GUID].GetString() == "Коррекция") return false;
            //MessageBox.Show(changedParam.ParameterInfo.EditType.ToString() + "\nIsReadOnly = " + changedParam.IsReadOnly.ToString() + "\nType" + changedParam.GetType());
            //если параметр только для чтения, то редактирование доступно
            if (changedParam.ParameterInfo.EditType == ParameterEditType.ChangeFromUIDenied) return false;
            //MessageBox.Show(changedParam.ParameterInfo.Name.ToString());
            string error = "Ошибка!\nНе доступно для изменения!\nНажмите клавишу Esc!";
            Ошибка(error);
            return true;
        }

        public void СобытиеСоздания()
        {
            AddNumber();
            ReferenceObject current = Context.ReferenceObject;

            if (current.IsCopy)
            {
                //MessageBox.Show("Copy");
                //вызывается при копировании для восстановления начальных значений второстепенных параметров
                ForCopyClearParams(current);
            }
            Errand er = new Errand(current);
            //вызываем карточку журнала регистрации для её создания
            var reg = er.RegistryJournalRecord;
        }

        public void ИзменениеИсполнителя()
        {
            Объект исполнитель = ТекущийОбъект.СвязанныйОбъект["Исполнитель поручения"];
            ReferenceObject currentObject = Context.ReferenceObject;
            var uiContext = Context as UIMacroContext;
            //MessageBox.Show(uiContext.ChangedLink.ToString()+" == "+"Рабочие объекты");
            if (uiContext != null && uiContext.ChangedLink.ToString() == "Исполнитель поручения")
            {
                Объект пользователь = ТекущийОбъект.СвязанныйОбъект["Исполнитель поручения"];
                if (пользователь != null)
                {
                    var должность = ВыполнитьМакрос("6a4385ff-475a-4afb-bb85-965252abebdc", "ОбъектДолжность", исполнитель);
                    //MessageBox.Show(должность.ToString());
                    if (должность != null)
                    {
                        var подразделение = ВыполнитьМакрос("6a4385ff-475a-4afb-bb85-965252abebdc", "Подразделение", должность);

                        ТекущийОбъект.Изменить();
                        //ТекущийОбъект.СвязанныйОбъект["Должность исполнителя"] = должность;
                        ТекущийОбъект.СвязанныйОбъект["Подразделение Исполнителя поручения"] = подразделение;
                        ТекущийОбъект.Сохранить();
                    }
                }
            }
        }

        public void ИзменениеРуководителя()
        {
            ReferenceObject currentObject = Context.ReferenceObject;
            var uiContext = Context as UIMacroContext;
            //MessageBox.Show(uiContext.ChangedLink.ToString()+" == "+"Рабочие объекты");
            if (uiContext != null && uiContext.ChangedLink.ToString() == "Руководитель")
            {
                Объект пользователь = ТекущийОбъект.СвязанныйОбъект["Руководитель поручения"];
                if (пользователь != null)
                {

                    var должность = ВыполнитьМакрос("6a4385ff-475a-4afb-bb85-965252abebdc", "ОбъектДолжность", пользователь);
                    var подразделение = ВыполнитьМакрос("6a4385ff-475a-4afb-bb85-965252abebdc", "Подразделение", должность);

                    ТекущийОбъект.Изменить();
                    //ТекущийОбъект.СвязанныйОбъект["Должность руководителя"] = должность;
                    ТекущийОбъект.СвязанныйОбъект["Подразделение Руководителя поручения"] = подразделение;
                    ТекущийОбъект.Сохранить();
                }
            }
        }

        public void ИзменениеКонтролёра()
        {

        }

        #endregion

        #region Нажатия на кнопки
        public void СоздатьПоручениеИзКанцелярии()
        {
            ReferenceObject rkk = Context.ReferenceObject;
            CreateErrandFromCancelariaCard(rkk);
        }

        /// <summary>
        /// Создаёт поручение на основе РКК - Входящие
        /// </summary>
        public bool CreateErrandFromCancelariaCard(ReferenceObject rkk)
        {

            ReferenceObject position = null;
            Объект подразделение = null;
            ReferenceObject department = null;

            Объект исполнитель = ТекущийОбъект.СвязанныйОбъект["Ответственный - Группы и пользователи"];
            if (исполнитель != null)
            {
                var должность = ВыполнитьМакрос("6a4385ff-475a-4afb-bb85-965252abebdc", "ОбъектДолжность", исполнитель);
                //Guid должностьGuid = new Guid(должность.Параметр["Guid"].ToString());
                //position = referenceUsers.Find(должностьGuid);
                подразделение = ВыполнитьМакрос("6a4385ff-475a-4afb-bb85-965252abebdc", "Подразделение", должность);
                //Guid подразделениеGuid = new Guid(подразделение.Параметр["Guid"].ToString());
                department = (ReferenceObject)подразделение;
            }

            // поиск типа
            ClassObject classObject = ErrandControlReference.Errand_Class;
            // создание
            ReferenceObject newObject = ErrandControlReference.Reference.CreateReferenceObject(classObject);
            // установка связи 1:1 или N:1 ответственный - исполнитель
            newObject.SetLinkedObject(link_Checker_GUID, (ReferenceObject)исполнитель);
            // установка связи 1:1 или N:1 ответственный - исполнитель
            newObject.SetLinkedObject(link_Director_GUID, (ReferenceObject)исполнитель);

            // установка связи 1:1 или N:1 Подразделение Исполнителя
            //newObject.SetLinkedObject(link_DepartmentPerformer_GUID, department);
            //ЗаполнитьИсполнителя();
            newObject[param_ErrandDocNum_GUID].Value = "Входящая корреспонденция №" + rkk[param_RegNum_GUID].Value.ToString();
            newObject.AddLinkedObject(link_ToInputData_GUID, rkk);
            newObject[param_Category_GUID].Value = "Входящая корреспонденция";
            foreach (ReferenceObject item in rkk.GetObjects(link_ToDocuments_GUID))
            {
                newObject.AddLinkedObject(link_ToInputData_GUID, item);
            }
            //newObject.ApplyChanges();

            RefObj newRefObj = RefObj.CreateInstance(newObject, Context);
            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true); // Сохраняем изменения и закрываем диалог

            //ShowPropertyDialog(newRefObj);
            return ShowPropertyDialog(newRefObj);
        }

        public void НажатиеНаКнопкуОтправитьИсполнителю()//рассылка первоначальная
        {
            ReferenceObject current = Context.ReferenceObject;
            Errand errand = new Errand(current);
            if (StartControlProcess(errand))
            {
                var uiContext = Context as UIMacroContext;
                uiContext.CloseDialog(true);
            }
        }

        public void НажатиеНаКнопкуРуководителя_Выполнено()
        {
            ReferenceObject current = Context.ReferenceObject;
            Errand errand = new Errand(current);
            ValidateWithResult(errand, Errand.ValidationResultValue.Done);

            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true);

        }

        public void НажатиеНаКнопкуРуководителя_НеВыполнено()
        {
            ReferenceObject current = Context.ReferenceObject;
            Errand errand = new Errand(current);
            ValidateWithResult(errand, Errand.ValidationResultValue.NotDone);

            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true);
        }


        public void НажатиеНаКнопкуРуководителя_Коррекция()
        {
            ReferenceObject current = Context.ReferenceObject;
            Errand errand = new Errand(current);
            StartCorrection(errand);

            if (!ReferenceObjectEditManager.Instance.IsEditingInDialog(current))
            {
                ShowPropertyDialog(CurrentObject);
            }
        }

        public void НажатиеНаКнопкуРуководителя_ЗакончитьКоррекцию()
        {
            ReferenceObject current = Context.ReferenceObject;
            Errand errand = new Errand(current);
            EndTheCorrection(errand);
            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true);
        }

        public void НажатиеНаКнопкуРуководителя_ОтменитьКоррекцию()
        {
            ReferenceObject current = Context.ReferenceObject;
            Errand errand = new Errand(current);
            CancelCorrection(errand);
            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true);
        }

        public void НажатиеНаКнопкуРуководителя_Аннулировать()
        {
            ReferenceObject current = Context.ReferenceObject;
            Errand errand = new Errand(current);
            ValidateWithResult(errand, Errand.ValidationResultValue.Annulated);

            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true);
        }


        public void НажатиеНаКнопкуКонтролёра_Выполнено()
        {
            ReferenceObject current = Context.ReferenceObject;
            Errand errand = new Errand(current);
            VerificateWithResult(errand, Errand.VerificationResultValue.Match);

            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true);
        }

        public void НажатиеНаКнопкуКонтролёра_НеВыполнено()
        {
            ReferenceObject current = Context.ReferenceObject;
            Errand errand = new Errand(current);
            VerificateWithResult(errand, Errand.VerificationResultValue.DoesNotMatch);

            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true);
        }

        public void НажатиеНаКнопкуКонтролёра_ТребуетсяКоррекция()
        {
            ReferenceObject current = Context.ReferenceObject;
            Errand errand = new Errand(current);
            VerificateWithResult(errand, Errand.VerificationResultValue.NeedCorrection);

            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true);
        }


        public void НажатиеНаКнопкуКонтролёра_ВернутьНаДоработку()//доработано
        {
            ReferenceObject current = Context.ReferenceObject;
            Errand errand = new Errand(current);
            ReturnToRework(errand);

            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true);
        }


        public void НажатиеНаКнопкуИсполнителя_Выполнено()
        {
            ReferenceObject current = Context.ReferenceObject;
            Errand errand = new Errand(current);
            PerformerResultChange(errand, Errand.PerformerResultValue.Done);

            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true);
        }

        public void НажатиеНаКнопкуИсполнителя_ИзменитьКонтрольнуюДату()
        {
            ReferenceObject current = Context.ReferenceObject;
            Errand errand = new Errand(current);
            ChangeNextCheckDate(errand);

            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true);
        }


        public void НажатиеНаКнопкуИсполнителя_ТребуетсяКоррекция()
        {

            ReferenceObject current = Context.ReferenceObject;
            Errand errand = new Errand(current);
            PerformerResultChange(errand, Errand.PerformerResultValue.NeedCorrection);

            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true);
        }

        public void НажатиеНаКнопкуИсполнителя_ПринятьВРаботу()
        {

            ReferenceObject current = Context.ReferenceObject;
            Errand errand = new Errand(current);
            PerformerResultChange(errand, Errand.PerformerResultValue.NotDone);

            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true);
        }

        public void НажатиеНаКнопкуИсполнителя_СделатьЗапись()
        {
            ReferenceObject current = Context.ReferenceObject;
            Errand errand = new Errand(current);
            AddToLogWithDialogShow(errand, "Промежуточная запись.", Errand.Role.Performer);

            var uiContext = Context as UIMacroContext;
            uiContext.CloseDialog(true);
        }

        public void НажатиеНаКнопкуАдминистратора_СделатьЗапись()
        {
            ReferenceObject current = Context.ReferenceObject;
            Errand errand = new Errand(current);
            AddToLogWithDialogShow(errand, "Администратор.", Errand.Role.None);
        }



        #region Административные команды
        public void Ручное_Обновление_Должностей()
        {
            ReferenceObject currentObject = Context.ReferenceObject;
            if (ТекущийОбъект.СвязанныйОбъект["Должность Контролёра"] != null &&
                ТекущийОбъект.СвязанныйОбъект["Должность исполнителя"] != null &&
                ТекущийОбъект.СвязанныйОбъект["Должность руководителя"] != null)
                return;

            Объект Контролёр = ТекущийОбъект.СвязанныйОбъект["Контролёр поручения"];
            Объект Исполнитель = ТекущийОбъект.СвязанныйОбъект["Исполнитель поручения"];
            Объект Руководитель = ТекущийОбъект.СвязанныйОбъект["Руководитель поручения"];
            Объект должностьКонтролёр = null;
            Объект должностьИсполнитель = null;
            Объект должностьРуководитель = null;

            if (Контролёр != null) должностьКонтролёр = ВыполнитьМакрос("6a4385ff-475a-4afb-bb85-965252abebdc", "ОбъектДолжность", Контролёр);
            if (Исполнитель != null) должностьИсполнитель = ВыполнитьМакрос("6a4385ff-475a-4afb-bb85-965252abebdc", "ОбъектДолжность", Исполнитель);
            if (Руководитель != null) должностьРуководитель = ВыполнитьМакрос("6a4385ff-475a-4afb-bb85-965252abebdc", "ОбъектДолжность", Руководитель);
            //подразделение = ВыполнитьМакрос("6a4385ff-475a-4afb-bb85-965252abebdc", "Подразделение", должность);

            if (должностьКонтролёр == null && должностьИсполнитель == null && должностьРуководитель == null) return;

            ТекущийОбъект.Изменить();
            if (должностьКонтролёр != null && ТекущийОбъект.СвязанныйОбъект["Должность Контролёра"] == null)
                ТекущийОбъект.СвязанныйОбъект["Должность Контролёра"] = должностьКонтролёр;
            if (должностьИсполнитель != null && ТекущийОбъект.СвязанныйОбъект["Должность исполнителя"] == null)
                ТекущийОбъект.СвязанныйОбъект["Должность исполнителя"] = должностьИсполнитель;
            if (должностьРуководитель != null && ТекущийОбъект.СвязанныйОбъект["Должность руководителя"] == null)
                ТекущийОбъект.СвязанныйОбъект["Должность руководителя"] = должностьРуководитель;
            ТекущийОбъект.Сохранить();

        }

        public void ChangePerformerResult()
        {
            ReferenceObject current = Context.ReferenceObject;
            ДиалогВвода диалог = СоздатьДиалогВвода("Введите значения");
            //List<string> values = new List<string>();
            var listValues = ErrandControlReference.Reference.ParameterGroup[param_PerformerResult_GUID].ValueList;
            object[] values = new object[listValues.Count];
            for (int i = 0; i < listValues.Count; i++)
            {
                values[i] = listValues[i].ToString();
            }
            диалог.ДобавитьВыборИзСписка("Выберите результат", values);
            диалог.ДобавитьСтроковое("Введите комментарий", "", true);
            Errand errand = new Errand(current);
            string oldPerformerResult = errand.PerformerResult;
            string newPerformerResult = "";
            string comment = "";
            диалог["Выберите результат"] = oldPerformerResult;
            if (диалог.Показать())
            {
                newPerformerResult = диалог["Выберите результат"];
                comment = диалог["Введите комментарий"];
                if (errand.PerformerResult != newPerformerResult)
                {
                    errand.PerformerResult = newPerformerResult;


                    comment = "Изменено значение валидации с " + oldPerformerResult + ", на " + newPerformerResult + ". Причина: " + comment;
                    //ДобавитьКомментарий(current, "Корректировка данных", (int)Errand.Role.None, false, comment);
                    AddRecordToLog(errand, DateTime.Now, "Корректировка данных", Errand.Role.None, false, comment);
                    errand.SaveSecondaryParametersToDB();
                }
            }
        }

        public void ChangeDirectorResult()
        {
            ReferenceObject current = Context.ReferenceObject;
            ДиалогВвода диалог = СоздатьДиалогВвода("Введите значения");
            var listValues = ErrandControlReference.Reference.ParameterGroup[param_ValidationResult_GUID].ValueList;
            object[] values = new object[listValues.Count];
            for (int i = 0; i < listValues.Count; i++)
            {
                values[i] = listValues[i].ToString();
            }
            диалог.ДобавитьВыборИзСписка("Выберите результат", values);
            диалог.ДобавитьСтроковое("Введите комментарий", "", true);
            Errand errand = new ErrandControl.Errand(current);
            string oldValidationResult = errand.ValidationResult;
            string newValidationResult = "";
            string comment = "";
            диалог["Выберите результат"] = oldValidationResult;
            if (диалог.Показать())
            {
                newValidationResult = диалог["Выберите результат"];
                comment = диалог["Введите комментарий"];
                if (oldValidationResult != newValidationResult)
                {
                    comment = "Изменено значение валидации с " + oldValidationResult + ", на " + newValidationResult + ". Причина: " + comment;
                    errand.ValidationResult = newValidationResult;

                    AddRecordToLog(errand, DateTime.Now, "Корректировка данных", Errand.Role.None, false, comment);
                    errand.SaveSecondaryParametersToDB();
                    //ДобавитьКомментарий(current, "Корректировка данных", (int)Errand.Role.None, false, comment);
                }
            }

        }

        #endregion

        #endregion
        #endregion
#endif



        public ServerConnection Connection
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

        private static Guid reference_Users_Guid = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");    //Guid справочника - "Группы и пользователи"
                                                                                                        //Справочник - "Запросы коррекции"
        private readonly Guid CorrectionRequests_reference_GUID = new Guid("d28ef384-b7b1-42e2-a1b2-37cddd826b5d");//Guid справочника - "Запросы коррекции"
        private readonly Guid CR_class_CorrectionRequest_GUID = new Guid("ba41c554-028e-4550-8179-d64b71266eea");//Guid типа "Запрос коррекции"
        private readonly Guid CR_list_RequisiteChanges_GUID = new Guid("195eba3e-4024-4887-8b22-8a59c5c50c8d");//Guid списка "Изменения реквизитов"
        private readonly Guid CR_param_ErrandGuid_GUID = new Guid("29ac358c-965a-4a80-8470-45c18a71ed97");//Guid параметра "Guid поручения"
        private readonly Guid CR_param_Reason_GUID = new Guid("ec4645d9-35c5-4b47-8917-2650c0764b39");//Guid параметра "Причина коррекции"
        private readonly Guid CR_param_ApplyDate_GUID = new Guid("7486b2c5-75cf-4c0d-8307-f0897dfe5bae");//Guid параметра "Дата применения изменения"
        private readonly Guid Users_КиселёваМарина_item_GUID = new Guid("ef037118-95ec-4346-ba6e-684301a4d875"); //Guid объекта справочника - "Группы и пользователи"

        private Reference _CorrectionRequestsReference;
        /// <summary>
        /// Справочник "Запросы коррекции"
        /// </summary>
        private Reference CorrectionRequestsReference
        {
            get
            {
                if (_CorrectionRequestsReference == null)

                    return GetReference(ref _CorrectionRequestsReference, CorrectionRequestsReferenceInfo);

                return _CorrectionRequestsReference;
            }
        }

        private ReferenceInfo _CorrectionRequestsReferenceInfo;

        private ReferenceInfo CorrectionRequestsReferenceInfo
        {
            get { return GetReferenceInfo(ref _CorrectionRequestsReferenceInfo, CorrectionRequests_reference_GUID); }
        }
        private ReferenceInfo UserReferenceInfo
        {
            get { return GetReferenceInfo(ref _userReferenceInfo, reference_Users_Guid); }
        }

        private ReferenceObject _КиселёваМарина;
        private ReferenceObject КиселёваМарина
        {
            get
            {
                if (_КиселёваМарина == null)
                    _КиселёваМарина = UserReference.Find(Users_КиселёваМарина_item_GUID);
                //return (ReferenceObject)CurrentUser;
                return _КиселёваМарина;
            }
        }




        #endregion


        private ReferenceInfo GetReferenceInfo(ref ReferenceInfo referenceInfo, Guid referenceGuid)
        {
            if (referenceInfo == null)
                referenceInfo = Connection.ReferenceCatalog.Find(referenceGuid);

            return referenceInfo;
        }

        string CurrentUserName
        {
            get
            {
                return ErrandControlReference.Connection.ClientView.UserName;
            }
        }

        private UserReference _userReference;
        /// <summary>
        /// Справочник "Группы и пользователи"
        /// </summary>
        private UserReference UserReference
        {
            get
            {
                if (_userReference == null)

                    _userReference = new UserReference(Connection);

                return _userReference;
            }
        }
        private ReferenceInfo _userReferenceInfo;


        private readonly Guid Users_Островой_item_GUID = new Guid("e86c95d6-dadd-4839-b5ce-efea2e88b1a3"); //Guid объекта справочника - "Группы и пользователи"
        private ReferenceObject _Островой;
        private ReferenceObject Островой
        {
            get
            {
                if (_Островой == null)
                    _Островой = UserReference.Find(Users_Островой_item_GUID);
                //return (ReferenceObject)CurrentUser;
                return _Островой;
            }
        }

        private readonly Guid Users_ВахромоваЯна_item_GUID = new Guid("9d2e9bec-102a-4c61-8810-994614835e5d"); //Guid объекта справочника - "Группы и пользователи"
        private ReferenceObject _ВахромоваЯна;
        private ReferenceObject ВахромоваЯна
        {
            get
            {
                if (_ВахромоваЯна == null)
                    _ВахромоваЯна = UserReference.Find(Users_ВахромоваЯна_item_GUID);
                //return (ReferenceObject)CurrentUser;
                return _ВахромоваЯна;
            }
        }

        private readonly Guid Users_СавченковаМария_item_GUID = new Guid("8806843b-19cf-4037-8228-e9728e42f9e5"); //Guid объекта справочника - "Группы и пользователи"
        private ReferenceObject _СавченковаМария;
        private ReferenceObject СавченковаМария
        {
            get
            {
                if (_СавченковаМария == null)
                    _СавченковаМария = UserReference.Find(Users_СавченковаМария_item_GUID);
                //return (ReferenceObject)CurrentUser;
                return _СавченковаМария;
            }
        }
    }

    interface IRegistryJournal
    {
        ErrandRegistryJournalRecord RegistryJournalRecord { get; }
    }
    internal class Errand : IRegistryJournal
    {

        bool beginChangesApplied;
        bool haveChangesToSave;
        public Errand(ReferenceObject errand_ro)
        {
            ReferenceObject = errand_ro;
            PerformerDoneDate = (DateTime?)ReferenceObject[EC_param_DoneTaskDate_GUID].Value;
            PerformerAcceptDate = (DateTime?)ReferenceObject[EC_param_AcceptTaskDate_GUID].Value;
            SendToPerformerDate = (DateTime?)ReferenceObject[EC_param_SendTaskDate_GUID].Value;
            ValidationDate = (DateTime?)ReferenceObject[EC_param_ValidationDate_GUID].Value;
            VerificationDate = (DateTime?)ReferenceObject[EC_param_VerificationDate_GUID].Value;
            NextCheckDate = ReferenceObject[EC_param_NextCheckDate_GUID].GetDateTime();
            beginChangesApplied = false;
            haveChangesToSave = false;
        }

        public string Stage
        {
            get
            {
                return ReferenceObject.SystemFields.Stage.ToString();
            }
        }

        public void ChangeStageTo(string stageName)
        {
            string methodName = "";
            switch (stageName)
            {
                case ErrandStages.Closed:
                    methodName = "ChangeStageClosed";
                    break;
                case ErrandStages.InWork:
                    methodName = "ChangeStageInWork";
                    break;
                default: break;
            }
            MacroContext mp = new MacroContext(ReferenceObject);
            mp.RunMacro("Запуск макроса на стороне сервера", "CreateServerTask", new object[] { "Контроль поручений.2.1 Сервер", methodName });
        }

        public Errand()
        {
        }

        ErrandRegistryJournalRecord _JournalRecord;
        public ErrandRegistryJournalRecord RegistryJournalRecord
        {
            get
            {
                if (_JournalRecord != null) return _JournalRecord;
                if (ReferenceObject != null)
                {
                    _JournalRecord = RegistryJournalRecordFinder.FindByGuid(this.ReferenceObject.SystemFields.Guid);
                    if (_JournalRecord == null)
                    {
                        _JournalRecord = new ErrandRegistryJournalRecord(this.ReferenceObject.SystemFields.Guid);
                        _JournalRecord.SaveToDB();
                        ReferenceObject.BeginChanges();
                        ReferenceObject.SetLinkedObject(EC_link_RegistryJournalRecord_Guid, _JournalRecord.ReferenceObject);
                        ReferenceObject.EndChanges();
                    }
                }
                return _JournalRecord;
            }
            private set
            {
                if (_JournalRecord == value) return;
                else _JournalRecord = value;
            }
        }

        public void SaveRequisitesToDB()
        {
            if (ReferenceObject == null)
            {
                ReferenceObject = CreateNewErrandInDB();

            }
            else
            {
                ReferenceObject refO;
                if (!ReferenceObject.TryBeginChanges(out refO)) throw new Exception_CancelAllActivities("Объект заблокирован другим пользователем.");
            }
            ReferenceObject[EC_param_ErrandText_GUID].Value = Text;
            ReferenceObject[EC_param_ErrandPlanEndDate_GUID].Value = PlanEndDate;
            ReferenceObject[EC_param_DocumentNumberWithErrand_GUID].Value = DocumentForErrand;
            ReferenceObject[EC_param_ProposedResult_GUID].Value = ProposedResult;
            ReferenceObject[EC_param_ExtPerformer_GUID].Value = ExternalPerformer;
            ReferenceObject[EC_param_ExtErrand_GUID].Value = IsPerformerExternal;
            ReferenceObject[EC_param_ExternalPerformerEmail_GUID].Value = ExternalPerformerEMail;
            ReferenceObject[EC_param_CoefficientOfImportance_GUID].Value = CoefficientOfImportance;
            this.ReferenceObject.SetLinkedObject(Errand.EC_link_Director_GUID, this.Director);
            this.ReferenceObject.SetLinkedObject(Errand.EC_link_Performer_GUID, this.Performer);
            this.ReferenceObject.SetLinkedObject(Errand.EC_link_Checker_GUID, this.Checker);

            ReferenceObject.EndChanges();
            beginChangesApplied = false;
            haveChangesToSave = false;
        }

        public void SaveSecondaryParametersToDB()
        {
            if (ReferenceObject == null)
            {
                ReferenceObject = CreateNewErrandInDB();

            }
            else
            {
                ReferenceObject refO;
                if (!ReferenceObject.TryBeginChanges(out refO)) throw new Exception_CancelAllActivities("Объект заблокирован другим пользователем.");
            }
            ReferenceObject[EC_param_CheckerComment_GUID].Value = CheckerComment;
            ReferenceObject[EC_param_DirectorComment_GUID].Value = DirectorComment;
            ReferenceObject[EC_param_PerformerComment_GUID].Value = PerformerComment;

            ReferenceObject[EC_param_DoneTaskDate_GUID].Value = PerformerDoneDate;
            ReferenceObject[EC_param_AcceptTaskDate_GUID].Value = PerformerAcceptDate;
            ReferenceObject[EC_param_SendTaskDate_GUID].Value = SendToPerformerDate;

            ReferenceObject[EC_param_NextCheckDate_GUID].Value = NextCheckDate;
            ReferenceObject[EC_param_VerificationDate_GUID].Value = VerificationDate;
            ReferenceObject[EC_param_ValidationDate_GUID].Value = ValidationDate;

            //результаты
            ReferenceObject[EC_param_PerformerResult_GUID].Value = PerformerResult;
            ReferenceObject[EC_param_VerificationResult_GUID].Value = VerificationResult;
            ReferenceObject[EC_param_ValidationResult_GUID].Value = ValidationResult;

            //лог
            ReferenceObject[EC_param_ErrandLog_GUID].Value = Log;

            ReferenceObject.EndChanges();
            beginChangesApplied = false;
            haveChangesToSave = false;
        }

        private ReferenceObject CreateNewErrandInDB()
        {
            ReferenceObject newErrand = ErrandControlReference.Reference.CreateReferenceObject(ErrandControlReference.Errand_Class);
            newErrand[EC_param_ErrandText_GUID].Value = Text;
            foreach (var item in GetInitialData())
            {
                newErrand.AddLinkedObject(EC_link_ToInputData_GUID, item);
            }
            newErrand[EC_param_DocumentNumberWithErrand_GUID].Value = DocumentForErrand;
            return newErrand;
        }

        bool? _IsPerformerExternal;
        public bool IsPerformerExternal
        {
            get
            {
                if (_IsPerformerExternal == null)
                    _IsPerformerExternal = ReferenceObject[EC_param_ExtErrand_GUID].GetBoolean();
                return (bool)_IsPerformerExternal;
            }
            set
            {
                if (_IsPerformerExternal != value)
                    _IsPerformerExternal = value;
            }

        }

        string _ExternalPerformer;
        public string ExternalPerformer
        {
            get
            {
                if (_ExternalPerformer == null && this.IsPerformerExternal)
                {
                    _ExternalPerformer = ReferenceObject[EC_param_ExtPerformer_GUID].GetString();
                }
                return _ExternalPerformer;
            }
            set
            {
                if (_ExternalPerformer != value)
                    _ExternalPerformer = value;
            }
        }

        string _ExternalPerformerEMail;
        public string ExternalPerformerEMail
        {
            get
            {
                if (_ExternalPerformerEMail == null && this.IsPerformerExternal)
                {
                    _ExternalPerformerEMail = ReferenceObject[EC_param_ExternalPerformerEmail_GUID].GetString();
                }
                return _ExternalPerformerEMail;
            }
            set
            {
                if (_ExternalPerformerEMail != value)
                    _ExternalPerformerEMail = value;
            }
        }

        User _Performer;

        public User Performer
        {
            get
            {
                if (_Performer == null && !this.IsPerformerExternal)
                    _Performer = ReferenceObject.GetObject(EC_link_Performer_GUID) as User;
                return _Performer;
            }
            set
            {
                if (_Performer != value)
                    _Performer = value;
            }
        }

        User _Checker;

        public User Checker
        {
            get
            {
                if (_Checker == null)
                    _Checker = ReferenceObject.GetObject(EC_link_Checker_GUID) as User;
                return _Checker;
            }
            set
            {
                if (_Checker != value)
                    _Checker = value;
            }
        }

        User _Director;

        public User Director
        {
            get
            {
                if (_Director == null)
                    _Director = ReferenceObject.GetObject(EC_link_Director_GUID) as User;
                return _Director;
            }
            set
            {
                if (_Director != value)
                    _Director = value;
            }
        }

        string _PerformerResult;
        public string PerformerResult
        {
            get
            {
                if (_PerformerResult == null && ReferenceObject != null) _PerformerResult = ReferenceObject[EC_param_PerformerResult_GUID].GetString();
                return _PerformerResult;
            }
            set
            {
                if (_PerformerResult != value)
                    _PerformerResult = value;
            }
        }

        string _PerformerComment;
        public string PerformerComment
        {
            get
            {
                if (_PerformerComment == null && ReferenceObject != null) _PerformerComment = ReferenceObject[EC_param_PerformerComment_GUID].GetString();
                return _PerformerComment;
            }
            set
            {
                if (_PerformerComment != value) _PerformerComment = value;
            }
        }

        string _CheckerComment;
        public string CheckerComment
        {
            get
            {
                if (_CheckerComment == null && ReferenceObject != null) _CheckerComment = ReferenceObject[EC_param_CheckerComment_GUID].GetString();
                return _CheckerComment;
            }
            set
            {
                if (_CheckerComment != value) _CheckerComment = value;
            }
        }

        string _DirectorComment;
        public string DirectorComment
        {
            get
            {
                if (_DirectorComment == null && ReferenceObject != null) _DirectorComment = ReferenceObject[EC_param_DirectorComment_GUID].GetString();
                return _DirectorComment;
            }
            set
            {
                if (_DirectorComment != value) _DirectorComment = value;
            }
        }

        string _VerificationResult;
        public string VerificationResult
        {
            get
            {
                if (_VerificationResult == null && ReferenceObject != null) _VerificationResult = ReferenceObject[EC_param_VerificationResult_GUID].GetString();
                return _VerificationResult;
            }
            set { if (_VerificationResult != value) _VerificationResult = value; }
        }

        string _ValidationResult;
        public string ValidationResult
        {
            get
            {
                if (_ValidationResult == null && ReferenceObject != null) _ValidationResult = ReferenceObject[EC_param_ValidationResult_GUID].GetString();
                return _ValidationResult;
            }
            set { if (_ValidationResult != value) _ValidationResult = value; }
        }



        string _Log;
        public string Log
        {
            get
            {
                if (_Log == null && ReferenceObject != null) _Log = ReferenceObject[EC_param_ErrandLog_GUID].GetString();
                return _Log;
            }

            set
            {
                if (_Log != value)
                    _Log = value;
            }
        }

        CorrectionRequest _CorrectionRequest;
        public CorrectionRequest ActualCorrectionRequest
        {
            get
            {
                if (_CorrectionRequest == null)
                    _CorrectionRequest = TryToFindExistedNotClosedCorrectionRequestFor(this);
                return _CorrectionRequest;
            }
            set
            {
                if (_CorrectionRequest != value)
                    _CorrectionRequest = value;
            }
        }

        private CorrectionRequest TryToFindExistedNotClosedCorrectionRequestFor(Errand errand)
        {
            CorrectionRequest cr = new CorrectionRequest(errand);
            if (cr.IsNew) cr = null;
            else if (cr.ApplyDate != null) cr = null;
            return cr;
        }

        int? _CoefficientOfImportance;
        public int CoefficientOfImportance
        {
            get
            {
                if (_CoefficientOfImportance == null && ReferenceObject != null) _CoefficientOfImportance = ReferenceObject[EC_param_CoefficientOfImportance_GUID].GetInt16();
                return (int)_CoefficientOfImportance;
            }
            set
            {
                if (_CoefficientOfImportance != value)
                    _CoefficientOfImportance = value;
            }
        }

        DateTime? _PlanEndDate;
        public DateTime PlanEndDate
        {
            get
            {
                if (_PlanEndDate == null)
                    _PlanEndDate = ReferenceObject[EC_param_ErrandPlanEndDate_GUID].GetDateTime();
                return (DateTime)_PlanEndDate;
            }
            set
            {
                if (_PlanEndDate != value)
                    _PlanEndDate = value;
            }
        }

        public DateTime NextCheckDate
        {
            get;
            set;
        }

        //DateTime? _PerformerDoneDate;
        public DateTime? PerformerDoneDate
        {
            get;
            set;
        }

        //DateTime? _PerformerAcceptDate;
        public DateTime? PerformerAcceptDate
        {
            get;
            set;
        }

        //DateTime? _SendToPerformerDate;
        public DateTime? SendToPerformerDate
        {
            get;
            set;
        }

        //DateTime? _ValidationDate;
        public DateTime? ValidationDate
        {
            get;
            set;
        }

        //DateTime? _VerificationDoneDate;
        public DateTime? VerificationDate
        {
            get;
            set;
        }

        int NumberOverDueDays
        {
            get
            {
                if (PerformerDoneDate == null) return (PlanEndDate - DateTime.Today).Days;
                return (PlanEndDate - (DateTime)PerformerDoneDate).Days;
            }
        }
        int CalculateCoefficientOfExecution()
        {
            if (PerformerResult != "Выполнено" && PlanEndDate > DateTime.Now) return 0;
            int A = 0;
            if (PerformerResult == "Не выполнено" && PlanEndDate < DateTime.Now) A = -1;
            else if (PerformerResult == "Выполнено") A = 1;
            int B = CoefficientOfImportance;
            int C = 0;
            if (A > 0)
            {
                if (NumberOverDueDays >= 30) C = 0;
                else if (NumberOverDueDays < 30 && NumberOverDueDays >= 2) C = 1;
                else if (NumberOverDueDays <= 1) C = 5;
            }
            if (A < 0)
            {
                if (NumberOverDueDays >= 30) C = 15;
                else if (NumberOverDueDays < 30 && NumberOverDueDays >= 2) C = 9;
                else if (NumberOverDueDays <= 1) C = 1;
            }
            int D = 0;
            switch (VerificationResult)
            {
                case "Соответствует":
                    D = 5;
                    break;
                default:
                    break;
            }
            int F = 0;
            switch (ValidationResult)
            {
                case "Выполнено":
                    F = 5;
                    break;
                default:
                    break;
            }

            //MessageBox.Show(string.Format("{0}*{1}({2}+{3}+{4})", A, B, C, D, F));

            return A * B * (C + D + F);
        }

        public int CoefficientOfExecution
        {
            get { return CalculateCoefficientOfExecution(); }
        }

        int _ErrandNumber;
        public int ErrandNumber
        {
            get
            {
                if (_ErrandNumber == 0)
                    _ErrandNumber = ReferenceObject[EC_param_ErrandNumber_GUID].GetInt16();
                return _ErrandNumber;
            }
        }

        string _Text;
        public string Text
        {
            get
            {
                if (_Text == null)
                {
                    if (ReferenceObject == null) _Text = "";
                    else _Text = ReferenceObject[EC_param_ErrandText_GUID].GetString();
                }
                return _Text;
            }
            set
            {
                if (value == _Text) return;
                _Text = value;
                if (ReferenceObject == null)
                {; }
                else
                {
                    if (!beginChangesApplied)
                    {
                        ReferenceObject.BeginChanges();
                        beginChangesApplied = true;
                    }
                    ReferenceObject[EC_param_ErrandText_GUID].Value = _Text;
                    haveChangesToSave = true;
                }
            }
        }

        List<ReferenceObject> _InitialData;

        public IEnumerable<ReferenceObject> GetInitialData()
        {
            if (_InitialData == null) _InitialData = new List<ReferenceObject>();
            return _InitialData;
        }

        public void AddInitialData(ReferenceObject initialData_ro)
        {
            if (_InitialData == null) _InitialData = new List<ReferenceObject>();
            _InitialData.Add(initialData_ro);
            if (ReferenceObject == null)
            {; }
            else
            {
                if (!beginChangesApplied)
                {
                    ReferenceObject.BeginChanges();
                    beginChangesApplied = true;
                }
                ReferenceObject.AddLinkedObject(EC_link_ToInputData_GUID, initialData_ro);
                haveChangesToSave = true;
            }
        }

        string _DocumentForErrand;
        public string DocumentForErrand
        {
            get
            {
                if (_DocumentForErrand == null)
                {
                    if (ReferenceObject == null) _DocumentForErrand = "";
                    else _DocumentForErrand = ReferenceObject[EC_param_DocumentNumberWithErrand_GUID].GetString();
                }
                return _DocumentForErrand;
            }
            set
            {
                if (value == _DocumentForErrand) return;
                _DocumentForErrand = value;
                if (ReferenceObject == null)
                {; }
                else
                {
                    if (!beginChangesApplied)
                    {
                        ReferenceObject.BeginChanges();
                        beginChangesApplied = true;
                    }
                    ReferenceObject[EC_param_DocumentNumberWithErrand_GUID].Value = _DocumentForErrand;
                    haveChangesToSave = true;
                }
            }
        }

        string _ProposedResult;
        public string ProposedResult
        {
            get
            {
                if (_ProposedResult == null && ReferenceObject != null) _ProposedResult = ReferenceObject[EC_param_ProposedResult_GUID].GetString();
                return _ProposedResult;
            }
            set
            {
                if (value == _ProposedResult) return;
                _ProposedResult = value;
                if (ReferenceObject == null)
                {; }
                else
                {
                    if (!beginChangesApplied)
                    {
                        ReferenceObject.BeginChanges();
                        beginChangesApplied = true;
                    }
                    ReferenceObject[EC_param_ProposedResult_GUID].Value = _ProposedResult;
                    haveChangesToSave = true;
                }
            }
        }


        public bool IsRequiredRequisitesFilled
        {
            get
            {
                bool allFieldsAssigned = true;
                //проверяем наличие текста поручения
                if (string.IsNullOrWhiteSpace(Text))
                    allFieldsAssigned = false;
                //проверяем наличие предполагаемого результата поручения
                if (allFieldsAssigned && string.IsNullOrWhiteSpace(ProposedResult))
                    allFieldsAssigned = false;
                //проверяем наличие предполагаемого результата поручения
                if (allFieldsAssigned && string.IsNullOrWhiteSpace(ProposedResult))
                    allFieldsAssigned = false;

                if (!allFieldsAssigned) System.Windows.Forms.MessageBox.Show("Не все реквизиты заполнены");
                return allFieldsAssigned;
            }
        }


        #region база данных
        public ReferenceObject ReferenceObject { get; set; }


        private static readonly Guid EC_param_DocumentNumberWithErrand_GUID = new Guid("f89f8ad5-7c33-4d68-8428-105b5991a3e9");//Guid параметра "Документ с поручением"
        private static readonly Guid EC_param_ErrandNumber_GUID = new Guid("c7e37e1d-2761-4527-a926-f4e54505d7cb");//Guid параметра "Номер поручения"
        private static readonly Guid EC_param_Category_GUID = new Guid("503cbb52-0204-4feb-9a17-cc83c5d4395e");//Guid параметра "Категория"
        private static readonly Guid EC_param_ErrandText_GUID = new Guid("dce7a86c-b2e5-4b32-935a-9420fab32558");//Guid параметра "Текст поручения"
        private static readonly Guid EC_param_ProposedResult_GUID = new Guid("d1cdf321-8cd2-493a-8747-9049038e0cfb");//Guid параметра "Ожидаемый результат"
        private static readonly Guid EC_param_ErrandLog_GUID = new Guid("72e07cfe-609d-49ce-9960-b451d71dc8e6");//Guid параметра "Журнал событий"
        private static readonly Guid EC_param_CoefficientOfImportance_GUID = new Guid("2c6d3089-6edc-427a-bf06-a69a24339a3a");//Guid параметра "Коэффициент важности"

        //Даты
        public static readonly Guid EC_param_ErrandPlanEndDate_GUID = new Guid("24bf25f7-bf81-405f-8f4e-75dda56e7ed3");//Guid параметра "Плановая дата выполнения поручения"
        private static readonly Guid EC_param_SendTaskDate_GUID = new Guid("39bef3f0-add4-4496-9c1f-a43a406253df");//Guid параметра "Дата отправки исполнителю"
        private static readonly Guid EC_param_AcceptTaskDate_GUID = new Guid("b82d7c2d-5568-43a5-9c0c-cdd049ff05e5");//Guid параметра "Дата принятия задания"
        private static readonly Guid EC_param_DoneTaskDate_GUID = new Guid("d4269274-e02a-4e92-ae92-ee470ebc756c");//Guid параметра "Дата выполнения исполнителем"
        private static readonly Guid EC_param_NextCheckDate_GUID = new Guid("681da3f7-c6ce-40f3-8e98-934a0fac444a");//Guid параметра "Дата следующего контроля"
        private static readonly Guid EC_param_VerificationDate_GUID = new Guid("07309e1e-4e97-492f-a2ad-303616988c40");//Guid параметра "Дата верификации"
        private static readonly Guid EC_param_ValidationDate_GUID = new Guid("878b3f6a-4e67-4e81-b073-0b50621ac46d");//Guid параметра "Дата валидации"

        //результаты
        private static readonly Guid EC_param_PerformerResult_GUID = new Guid("22124188-0f75-4417-961d-36a66cac0e7a");//Guid параметра "Результат исполнения"
        private static readonly Guid EC_param_VerificationResult_GUID = new Guid("8632c88c-fc67-4f1e-a650-72d11a35f2e0");//Guid параметра "Результат верификации"
        private static readonly Guid EC_param_ValidationResult_GUID = new Guid("7fc12479-9f81-44b5-a6f7-b74706c52b66");//Guid параметра "Результат валидации"

        //комментарии
        private static readonly Guid EC_param_DirectorComment_GUID = new Guid("8df3ad51-3dda-4045-a5c5-cb3ae157a486");//Guid параметра "Рецензия Руководителя"
        private static readonly Guid EC_param_PerformerComment_GUID = new Guid("2e0f9554-a5b7-4844-b1c6-422311e0869b");//Guid параметра "Комментарий исполнителя"
        private static readonly Guid EC_param_CheckerComment_GUID = new Guid("d5f428cf-9f17-4aaa-a4a1-964eff6713b4");//Guid параметра "Рецензия Контролёра"

        //роли и должности
        private static readonly Guid EC_link_Director_GUID = new Guid("587b07dc-fc72-43f8-9fb2-e698f52db3c3");//Guid связи n:1 "Руководитель"
        private static readonly Guid EC_param_ExtErrand_GUID = new Guid("5ed4ffce-30c7-4308-8a8c-1e2974af379e");//Guid параметра "Внешнее поручение"
        private static readonly Guid EC_param_ExtPerformer_GUID = new Guid("52129989-7ec0-4984-9c1a-77658ae1e763");//Guid параметра "Внешний исполнитель"
        private static readonly Guid EC_param_ExternalPerformerEmail_GUID = new Guid("6131602b-933a-4895-8aeb-542c216cfa1c");//Guid параметра "E-mail внешнего исполнителя"
        private static readonly Guid EC_link_Performer_GUID = new Guid("7cd90bfa-e37c-4265-a7e1-f65812940575");//Guid связи n:1 "Исполнитель поручения"
        private static readonly Guid EC_link_Checker_GUID = new Guid("10c6727f-f6a7-4550-a0ee-da862ff14d12");//Guid связи n:1 "Контролёр поручения"
        private static readonly Guid EC_link_PositionDirector_GUID = new Guid("830bdf4f-6e8c-4c83-9790-be29b3fe55be");//Guid связи n:1 "Должность Руководителя"
        private static readonly Guid EC_link_PositionPerformer_GUID = new Guid("8c18ac88-697a-4d22-afdc-49dd557b5cca");//Guid связи n:1 "Должность Исполнителя"
        private static readonly Guid EC_link_PositionChecker_GUID = new Guid("3508946b-d8d2-407c-8fe1-1205f158e794");//Guid связи n:1 "Должность Контролёра"
        private static readonly Guid EC_link_RegistryJournalRecord_Guid = new Guid("c2096bae-3a60-4269-b9f5-42e91ecc2ced");//Guid связи 1:1 "Журнал регистрации"


        private static readonly Guid EC_link_DepartmentPerformer_GUID = new Guid("a93bcf81-eb0b-4856-a905-8147e28c54e6");//Guid связи n:1 "Подразделение исполнителя"

        private static readonly Guid EC_link_ToInputData_GUID = new Guid("5b41779b-122e-4a53-8feb-a6388e44c81d");//Guid связи n:n "Исходные данные"

        #endregion

        #region перечисления и статические данные

        public enum Role { None, Performer, Checker, Director }
        public struct PerformerResultValue
        {
            public const string NotAccepted = "Не принято в работу";
            public const string Done = "Выполнено";
            public const string NotDone = "Не выполнено";
            public const string NeedCorrection = "Требуется коррекция";
        }
        public struct VerificationResultValue
        {
            public const string NotVerified = "Не верифицировано";
            public const string Match = "Соответствует";
            public const string DoesNotMatch = "Не соответствует";
            public const string NeedCorrection = "Требуется коррекция";
        }

        public struct ValidationResultValue
        {
            public const string NotValidated = "Не валидировано";
            public const string Done = "Выполнено";
            public const string NotDone = "Не выполнено";
            public const string Annulated = "Аннулировано";
            public const string Correction = "Коррекция";
        }

        public struct ErrandStages
        {
            public const string New = "Новый объект";
            public const string InWork = "В работе_Контроль поручений";
            public const string Closed = "Закрыто";
        }

        internal struct RequisiteNameForLog
        {
            internal const string Text = "<Текст поручения>";
            internal const string DocumentForErrand = "<Документ с поручением>";
            internal const string PlanEndDate = "<Срок исполнения>";
            internal const string ProposedResult = "<Ожидаемый результат>";
            internal const string Director = "<Руководитель>";
            internal const string Performer = "<Исполнитель>";
            internal const string IsExternalPerformer = "<Внешний исполнитель>";
            internal const string ExternalPerformerEMail = "<E-Mail внешнего исполнителя>";
            internal const string Checker = "<Контролёр>";
            internal const string ImportanceCoefficient = "<Коэффициент важности>";
        }

        public struct Requisites
        {

            public const string PlanDate = "Плановая дата выполнения";
            public const string Text = "Текст поручения";
            public const string Result = "Ожидаемый результат";
        }

        //словарь типов данных реквизитов
        public static readonly Dictionary<string, Type> RequisitesTypesDictionary = new Dictionary<string, Type>
        {
            { Requisites.PlanDate, typeof(DateTime) },
            { Requisites.Text, typeof(string) },
            { Requisites.Result, typeof(string) }
        };
        //словарь гуидов параметров реквизитов
        public static Dictionary<string, Guid> RequisitesGuidDictionary = new Dictionary<string, Guid>
        {
            { Requisites.PlanDate, EC_param_ErrandPlanEndDate_GUID },
            { Requisites.Text, EC_param_ErrandText_GUID },
            { Requisites.Result, EC_param_ProposedResult_GUID }
        };
        #endregion

        // Специальное исключение
        [Serializable]
        public class Exception_CancelAllActivities : ApplicationException
        {
            public Exception_CancelAllActivities() { }
            public Exception_CancelAllActivities(string message) : base(message) { }
            public Exception_CancelAllActivities(string message, Exception ex) : base(message) { }
            // Конструктор для обработки сериализации типа
            protected Exception_CancelAllActivities(System.Runtime.Serialization.SerializationInfo info,
                                                    System.Runtime.Serialization.StreamingContext contex)
                : base(info, contex)
            { }
        }
    }


    internal static class ErrandControlReference
    {

        public static ServerConnection Connection
        {
            get { return ServerGateway.Connection; }
        }



        private static Reference GetReference(ref Reference reference, ReferenceInfo referenceInfo)
        {
            if (reference == null)
                reference = referenceInfo.CreateReference();
            else reference.Refresh();

            return reference;
        }


        private static Reference _ErrandsControlReference;
        /// <summary>
        /// Справочник "Контроль поручений"
        /// </summary>
        internal static Reference Reference
        {
            get
            {
                if (_ErrandsControlReference == null)

                    return GetReference(ref _ErrandsControlReference, ErrandsControlReferenceInfo);

                return _ErrandsControlReference;
            }
        }

        private static ReferenceInfo _ErrandsControlReferenceInfo;

        private static ReferenceInfo ErrandsControlReferenceInfo
        {
            get { return GetReferenceInfo(ref _ErrandsControlReferenceInfo, ErrandControl_reference_GUID); }
        }

        private static ClassObject _ErrandClass;
        /// <summary>
        /// Тип - Поручение
        /// </summary>
        internal static ClassObject Errand_Class
        {
            get
            {
                if (_ErrandClass == null)
                    _ErrandClass = Reference.Classes.Find(EC_class_Errand_GUID);
                return _ErrandClass;
            }
        }


        private static ReferenceInfo GetReferenceInfo(ref ReferenceInfo referenceInfo, Guid referenceGuid)
        {
            if (referenceInfo == null)
                referenceInfo = Connection.ReferenceCatalog.Find(referenceGuid);

            return referenceInfo;
        }

        //Справочник - "Контроль поручений"
        private static readonly Guid ErrandControl_reference_GUID = new Guid("3a4e96a1-5f28-4d47-9fb4-fc2d48d0e98c");//Guid справочника - "Контроль поручений"
        private static readonly Guid EC_class_Errand_GUID = new Guid("31fd4174-6c0b-45db-a659-1174854421ee");//Guid типа "Поручение"
    }
    internal class CorrectionRequest
    {

        public override string ToString()
        {
            StringBuilder result = new StringBuilder();
            result.AppendLine("Причина: " + Reason);
            foreach (var item in RequisiteChangeNotices)
            {
                result.AppendLine(item.ToString());
            }
            return result.ToString();
        }


        List<RequisiteChangeNotice> _RequisiteChangeNotices;
        public List<RequisiteChangeNotice> RequisiteChangeNotices
        {
            get
            {
                if (_RequisiteChangeNotices == null && CorrectionRequest_ro != null)
                {
                    List<RequisiteChangeNotice> RequisiteChangeNoticesTemp = new List<RequisiteChangeNotice>();
                    foreach (var item in CorrectionRequest_ro.GetObjects(CR_list_RequisiteChanges_GUID))
                    {
                        RequisiteChangeNoticesTemp.Add(new RequisiteChangeNotice(item));
                    }
                    _RequisiteChangeNotices = RequisiteChangeNoticesTemp;
                }
                return _RequisiteChangeNotices;
            }
            private set
            {
                if (_RequisiteChangeNotices != value)
                    _RequisiteChangeNotices = value;
            }
        }

        public CorrectionRequest(Errand errand)
        {
            this.Errand = errand;
            ErrandGuid = errand.ReferenceObject.SystemFields.Guid.ToString();
            CorrectionRequest_ro = CorrectionRequestsReference.Find(CorrectionRequestsReferenceInfo.Description[CR_param_ErrandGuid_GUID], ErrandGuid).Where(cr => cr[CR_param_ApplyDate_GUID].IsNull).FirstOrDefault();
            if (CorrectionRequest_ro != null)
            {
                IsNew = false;
                this.Reason = CorrectionRequest_ro[CR_param_Reason_GUID].GetString();
                this.ApplyDate = (DateTime?)CorrectionRequest_ro[CR_param_ApplyDate_GUID].Value;
                this.Comment = CorrectionRequest_ro[CR_param_Comment_GUID].GetString();
            }
            else
            {
                this.Reason = errand.PerformerComment;
                this.RequisiteChangeNotices = new List<RequisiteChangeNotice>();
                IsNew = true;
            }
        }

        public CorrectionRequest(ReferenceObject cr_ro)
        {
            this.CorrectionRequest_ro = cr_ro;
            IsNew = false;
            this.ErrandGuid = CorrectionRequest_ro[CR_param_ErrandGuid_GUID].GetString();
            this.Reason = CorrectionRequest_ro[CR_param_Reason_GUID].GetString();
            this.ApplyDate = (DateTime?)CorrectionRequest_ro[CR_param_ApplyDate_GUID].Value;
            this.Comment = CorrectionRequest_ro[CR_param_Comment_GUID].GetString();
            //ApplyDate = (DateTime?)CorrectionRequest_ro[CR_param_ApplyDate_GUID].Value;
        }

        public void Apply(string comment)
        {
            Comment = comment;
            ApplyDate = DateTime.Now;
            SaveToDB();
        }
#if !server
        public bool ApproveByChecker()
        {
            if (!ShowDialog()) return false;
            foreach (var requisiteChangeNotice in this.RequisiteChangeNotices)
            {
                requisiteChangeNotice.ApproveByChecker();
            }
            SaveToDB();
            return true;
        }

        public bool ShowDialog()
        {
            UIMacroContext uIMacroContext = new UIMacroContext(new MacroContext(Connection), null, null, null, null);
            var dial = uIMacroContext.CreateInputDialog();
            dial.SetSize(400, 100);
            dial.Caption = "Выберите необходимые изменения";
            dial.AddStringField("Введите причину коррекции", this.Reason, true, true, "", true, 0);
            dial.AddGroup("Реквизиты");
            foreach (var item in Errand.RequisitesTypesDictionary)
            {
                RequisiteChangeNotice rcn = RequisiteChangeNotices.Where(r => r.ChangingParameterName == item.Key).FirstOrDefault();
                if (item.Value == typeof(DateTime))
                {
                    DateTime newValue;
                    if (rcn != null)
                    {
                        newValue = (DateTime)rcn.NewValue;
                    }
                    else newValue = Errand.ReferenceObject[Errand.RequisitesGuidDictionary[item.Key]].GetDateTime();
                    dial.AddDateField(item.Key, newValue, true, (int)DateTimePickerFormat.Short, "dd.MM.YYYY");
                    //if(newValue != Errand.)
                    dial.AddComment(item.Key, "");
                }
                else
                {
                    string newValue;
                    if (rcn != null)
                    {
                        newValue = rcn.NewValue;
                    }
                    else
                    {
                        newValue = Errand.ReferenceObject[Errand.RequisitesGuidDictionary[item.Key]].GetString();
                    }
                    dial.AddStringField(item.Key, newValue, true, true, "", true, 0);
                    dial.AddComment(item.Key, "");
                }
            }

            dial.FieldValueChanged += Dial_FieldValueChanged;
            //MessageBox.Show("dial - " + dial.GetType().ToString());
            if (dial.Show(uIMacroContext))
            {
                //MessageBox.Show("dial - " + (string)(dial.GetValue("Введите причину коррекции")));
                this.Reason = (string)(dial.GetValue("Введите причину коррекции"));
                //MessageBox.Show("dial - this.Reason -" + this.Reason);
                foreach (var errandRequisiteType in Errand.RequisitesTypesDictionary)
                {
                    if (dial.GetValue(errandRequisiteType.Key).ToString() != Errand.ReferenceObject[Errand.RequisitesGuidDictionary[errandRequisiteType.Key]].Value.ToString())
                    {
                        //MessageBox.Show("AddRequisiteChangeNotices - " + errandRequisiteType.Key);
                        AddRequisiteChangeNotices(errandRequisiteType.Key, dial.GetValue(errandRequisiteType.Key));
                    }
                }
                SaveToDB();
                return true;
            }
            return false;
        }

        private void Dial_FieldValueChanged(object sender, FieldValueChangedEventArgs e)
        {
            //ValuesDialog dial = sender as ValuesDialog;
            //dial.AddFlagProperty("изменение");
            //dial.AddComment(e.Name, "Изменено");
            //MessageBox.Show("sender" + sender.GetType().ToString());
            //MessageBox.Show("e" + e.ToString());
        }
#endif
        private void SaveToDB()
        {
            if (CorrectionRequest_ro == null)
            {
                CorrectionRequest_ro = CorrectionRequestsReference.CreateReferenceObject(CorrectionRequest_Class);
                CorrectionRequest_ro[CR_param_ErrandGuid_GUID].Value = Errand.ReferenceObject.SystemFields.Guid.ToString();
                this.CorrectionRequest_ro.EndChanges();
            }
            this.CorrectionRequest_ro.BeginChanges();
            //сохраняем изменения в списке изменений реквизитов( RequisiteChangeNotices)
            //MessageBox.Show("RequisiteChangeNotices.Count" + RequisiteChangeNotices.Count.ToString());
            //MessageBox.Show("this.Reason" + this.Reason);
            foreach (var requisiteChangeNotice in RequisiteChangeNotices)
            {
                requisiteChangeNotice.SaveToDB(this);
            }
            this.CorrectionRequest_ro[CR_param_ErrandGuid_GUID].Value = this.ErrandGuid;
            this.CorrectionRequest_ro[CR_param_Reason_GUID].Value = this.Reason;
            this.CorrectionRequest_ro[CR_param_Comment_GUID].Value = this.Comment;
            this.CorrectionRequest_ro[CR_param_ApplyDate_GUID].Value = this.ApplyDate;
            this.CorrectionRequest_ro.EndChanges();
        }

        private void AddRequisiteChangeNotices(string requisiteName, dynamic newValue)
        {
            RequisiteChangeNotice newRCN = RequisiteChangeNotices.Where(r => r.ChangingParameterName == requisiteName).FirstOrDefault();
            if (newRCN == null)
            {
                newRCN = CreateRequisiteChangeNotice(requisiteName, newValue);
                RequisiteChangeNotices.Add(newRCN);
            }
            else
            {
                newRCN.NewValue = newValue;
            }
        }

        Errand Errand { get; set; }

        public string Reason
        {
            get; private set;
        }

        public DateTime? ApplyDate
        {
            get; private set;
        }

        string _Comment;
        public string Comment
        {
            get
            {
                if (_Comment == null) return "";
                return _Comment;
            }
            private set
            {
                if (_Comment != value)
                    _Comment = value;
            }
        }

        public string ErrandGuid
        {
            get; private set;
        }

        //Создаёт новую коррекцию параметра
        private RequisiteChangeNotice CreateRequisiteChangeNotice(string paramName, dynamic newValue)
        {

            RequisiteChangeNotice newRequisiteChangeNotice = new RequisiteChangeNotice(paramName, Errand.RequisitesGuidDictionary[paramName].ToString(), newValue);
            return newRequisiteChangeNotice;
        }

        ReferenceObject _CorrectionRequest_ro;
        public ReferenceObject CorrectionRequest_ro
        {
            get
            {
                return _CorrectionRequest_ro;
            }
            private set
            {
                if (value != _CorrectionRequest_ro)
                    _CorrectionRequest_ro = value;
            }
        }
        #region база данных
        //Справочник - "Запросы коррекции"
        private readonly Guid CorrectionRequests_reference_GUID = new Guid("d28ef384-b7b1-42e2-a1b2-37cddd826b5d");//Guid справочника - "Запросы коррекции"
        private readonly Guid CR_class_CorrectionRequest_GUID = new Guid("ba41c554-028e-4550-8179-d64b71266eea");//Guid типа "Запрос коррекции"
        private readonly Guid CR_list_RequisiteChanges_GUID = new Guid("195eba3e-4024-4887-8b22-8a59c5c50c8d");//Guid списка "Изменения реквизитов"
        private readonly Guid CR_param_ErrandGuid_GUID = new Guid("29ac358c-965a-4a80-8470-45c18a71ed97");//Guid параметра "Guid поручения"
        private readonly Guid CR_param_Reason_GUID = new Guid("ec4645d9-35c5-4b47-8917-2650c0764b39");//Guid параметра "Причина коррекции"
        private readonly Guid CR_param_ApplyDate_GUID = new Guid("7486b2c5-75cf-4c0d-8307-f0897dfe5bae");//Guid параметра "Дата применения изменения"
        private readonly Guid CR_param_Comment_GUID = new Guid("1c241659-27db-458d-9190-e00ec5fb6802");//Guid параметра "Комментарий"


        //поля
        private Reference _CorrectionRequestsReference;
        private ReferenceInfo _CorrectionRequestsReferenceInfo;
        private ClassObject _CorrectionRequestClass;

        ServerConnection Connection
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
        /// Справочник "Запросы коррекции"
        /// </summary>
        private Reference CorrectionRequestsReference
        {
            get
            {
                if (_CorrectionRequestsReference == null)

                    return GetReference(ref _CorrectionRequestsReference, CorrectionRequestsReferenceInfo);

                return _CorrectionRequestsReference;
            }
        }

        #endregion
        private ReferenceInfo CorrectionRequestsReferenceInfo
        {
            get { return GetReferenceInfo(ref _CorrectionRequestsReferenceInfo, CorrectionRequests_reference_GUID); }
        }

        private ReferenceInfo GetReferenceInfo(ref ReferenceInfo referenceInfo, Guid referenceGuid)
        {
            if (referenceInfo == null)
                referenceInfo = Connection.ReferenceCatalog.Find(referenceGuid);

            return referenceInfo;
        }

        /// <summary>
        /// Тип - Запрос коррекции
        /// </summary>
        private ClassObject CorrectionRequest_Class
        {
            get
            {
                if (_CorrectionRequestClass == null)
                    _CorrectionRequestClass = CorrectionRequestsReference.Classes.Find(CR_class_CorrectionRequest_GUID);
                return _CorrectionRequestClass;
            }
        }

        public bool IsNew { get; private set; }

        #endregion
    }
    internal class RequisiteChangeNotice
    {
        CorrectionRequest CorrectionRequest;
        private ReferenceObject RequisiteChangeNotice_ro
        {
            get;
            set;

        }

        public RequisiteChangeNotice(CorrectionRequest correctionRequest)
        {
            CorrectionRequest = correctionRequest;
            //this.RequisiteChangeNotice_ro = requisiteChangeNotice_ro;
            this.ChangingParameterName = RequisiteChangeNotice_ro[RCN_param_ErrandParamName_GUID].GetString();
            this.ChangingParameterGuid = RequisiteChangeNotice_ro[RCN_param_ErrandParamGuid_GUID].GetString();

            if (Errand.RequisitesTypesDictionary[ChangingParameterName] == typeof(DateTime))
                this.NewValue = RequisiteChangeNotice_ro[RCN_param_NewDate_GUID].GetDateTime();
            else this.NewValue = RequisiteChangeNotice_ro[RCN_param_NewString_GUID].GetString();

            this.IsApproveByChecker = false;
            this.CheckerApproveDate = null;
        }

        public RequisiteChangeNotice(ReferenceObject requisiteChangeNotice_ro)
        {
            //CorrectionRequest = correctionRequest;
            this.RequisiteChangeNotice_ro = requisiteChangeNotice_ro;
            this.ChangingParameterName = RequisiteChangeNotice_ro[RCN_param_ErrandParamName_GUID].GetString();
            this.ChangingParameterGuid = RequisiteChangeNotice_ro[RCN_param_ErrandParamGuid_GUID].GetString();

            if (Errand.RequisitesTypesDictionary[ChangingParameterName] == typeof(DateTime))
                this.NewValue = RequisiteChangeNotice_ro[RCN_param_NewDate_GUID].GetDateTime();
            else this.NewValue = RequisiteChangeNotice_ro[RCN_param_NewString_GUID].GetString();

            this.IsApproveByChecker = false;
            this.CheckerApproveDate = null;
            //MessageBox.Show("RequisiteChangeNotice constr");
        }


        public RequisiteChangeNotice(string parameterName, string parameterGuid, dynamic newValue)
        {
            //CorrectionRequest = correctionRequest;
            //CorrectionRequest_ro = correctionRequest_ro;
            this.ChangingParameterName = parameterName;
            this.ChangingParameterGuid = parameterGuid;
            this.NewValue = newValue;
            this.IsApproveByChecker = false;
            this.CheckerApproveDate = null;
        }

        public void ApproveByChecker()
        {
            CheckerApproveDate = DateTime.Now;
            IsApproveByChecker = true;
        }

        public void SaveToDB(CorrectionRequest correctionRequest)
        {
            CorrectionRequest = correctionRequest;
            //MessageBox.Show(CorrectionRequest.CorrectionRequest_ro.SystemFields.Id.ToString());
            if (RequisiteChangeNotice_ro == null) this.RequisiteChangeNotice_ro = CorrectionRequest.CorrectionRequest_ro.CreateListObject(RCN_list_RequisiteChanges_GUID, RCN_class_ChangeRequisite_GUID);
            else RequisiteChangeNotice_ro.BeginChanges();
            RequisiteChangeNotice_ro.Reload();
            if (NewValue.GetType() == typeof(DateTime))
                RequisiteChangeNotice_ro[RCN_param_NewDate_GUID].Value = this.NewValue;
            else
                RequisiteChangeNotice_ro[RCN_param_NewString_GUID].Value = this.NewValue;

            RequisiteChangeNotice_ro[RCN_param_ErrandParamGuid_GUID].Value = this.ChangingParameterGuid;
            RequisiteChangeNotice_ro[RCN_param_ErrandParamName_GUID].Value = this.ChangingParameterName;
            RequisiteChangeNotice_ro[RCN_param_ApproveByChecker_GUID].Value = this.IsApproveByChecker;
            RequisiteChangeNotice_ro[RCN_param_ApproveByCheckerDate_GUID].Value = this.CheckerApproveDate;
            RequisiteChangeNotice_ro.EndChanges();
        }


        public override string ToString()
        {
            String result = String.Format("Изменение параметра: {0} - {1}", this.ChangingParameterName, this.NewValueToString);
            return result;
        }

        //string _ChangingParameterName;
        public string ChangingParameterName
        {
            get;
            private set;
        }

        string _ChangingParameterGuid;
        public string ChangingParameterGuid
        {
            get
            {
                if (_ChangingParameterGuid == null && RequisiteChangeNotice_ro != null)
                    _ChangingParameterGuid = RequisiteChangeNotice_ro[RCN_param_ErrandParamGuid_GUID].GetString();
                return _ChangingParameterGuid;
            }
            private set
            {
                if (_ChangingParameterGuid != value)
                    _ChangingParameterGuid = value;
            }
        }

        dynamic _NewValue;
        public dynamic NewValue
        {
            get
            {
                return _NewValue;
            }

            set
            {
                if (value != _NewValue)
                    _NewValue = value;
            }
        }

        public string NewValueToString
        {
            get
            {
                string result;
                if (NewValue is DateTime) result = ((DateTime)this.NewValue).ToShortDateString();
                else result = this.NewValue.ToString();
                return result;
            }
        }

        public bool IsApproveByChecker { get; private set; }
        public DateTime? CheckerApproveDate { get; private set; }

        //Тип "Изменение реквизитов"
        private readonly Guid RCN_list_RequisiteChanges_GUID = new Guid("195eba3e-4024-4887-8b22-8a59c5c50c8d");//Guid списка "Изменения реквизитов"
        private readonly Guid RCN_class_ChangeRequisite_GUID = new Guid("ec7e5b6a-36a6-40ef-a586-f7c14d32ce88");//Guid типа "Плановая дата выполнения"

        private readonly Guid RCN_param_ErrandParamGuid_GUID = new Guid("7d37d607-6934-4693-b23f-68eaa516c526");//Guid параметра "Guid параметра поручения"
        private readonly Guid RCN_param_ErrandParamName_GUID = new Guid("57bcea13-0442-4365-9700-2a5cf7b42bc9");//Guid параметра "Наименование параметра поручения"

        private readonly Guid RCN_param_NewDate_GUID = new Guid("ea2e51ce-cc4e-4de8-bd02-783688c7c4fa");//Guid параметра "Новая дата"
        private readonly Guid RCN_param_NewString_GUID = new Guid("7f124d32-7580-4dea-915f-77b12366cbd9");//Guid параметра "Новая строка"

        private readonly Guid RCN_param_ApproveByChecker_GUID = new Guid("0a4414f5-6480-4d6c-b2fc-e705c6b6ce15");//Guid параметра "Подтверждено контролёром"
        private readonly Guid RCN_param_ApproveByCheckerDate_GUID = new Guid("96fcdef1-c8b6-46e7-a1ba-946adcaa803c");//Guid параметра "Дата подтверждения контролёром"

    }

    /// <summary>
    /// HTML письмо
    /// ver 1.8
    /// </summary>
    class HTML_MailMessage
    {
        public HTML_MailMessage(string subject, string bodyText, List<ReferenceObject> mailRecipients, List<string> emailAdresses, List<ReferenceObject> mailAttachments)
        {
            Subject = subject;
            BodyText = bodyText;
            MailAttachments = mailAttachments ?? new List<ReferenceObject>();
            MailRecipients = mailRecipients ?? new List<ReferenceObject>();
            EmailAdresses = emailAdresses;
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
        public List<ReferenceObject> MailRecipients { get; private set; }
        public List<string> EmailAdresses { get; private set; }
        public void Send()
        {
            MailMessage = new MailMessage(Connection.Mail.DOCsAccount);
            MailMessage.Subject = CleanString(Subject);
            string messageBody = CreateMessageBody();
            MailMessage.SetBody(messageBody, MailBodyType.Html);
            FillAttachments();
            FillLocalFiles();
            FillMailRecipients();
            FillEmailAdresses();
            if (MailMessage.To == null || MailMessage.To.Count == 0) return;
            else
                MailMessage.Send();
        }

        private void FillEmailAdresses()
        {
            if (EmailAdresses == null) return;
            foreach (var adress in EmailAdresses)
            {
                if (!string.IsNullOrWhiteSpace(adress))
                {
                    EMailAddress eMailAdress = new EMailAddress(adress);
                    MailMessage.To.Add(eMailAdress);
                }
            }
        }

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

                if (!mailuser.Email.IsNull && !mailuser.Email.IsEmpty)
                {
                    MailMessage.To.Add(new EMailAddress(mailuser.Email.ToString()));
                }
            }
        }

        private void FillAttachments()
        {
            foreach (var attachment in MailAttachments)
            {
                MailMessage.Attachments.Add(new ObjectAttachment(attachment));
            }
        }

        MailMessage MailMessage;
        string BodyHeaderForSystem = @"<html>
	<head>
		<meta charset=""utf-8"">
	</head>
    <body><span style='font-size:12.0pt; font-family:""Arial"",""sans-serif"";mso-fareast-font-family:""Times New Roman"";
mso-fareast-theme-font:minor-fareast;mso-fareast-language:RU;mso-no-proof:yes'
    <p>Здравствуйте!</p>
    <p>Это автоматически созданное письмо, не отвечайте на него.</p>";

        string BodyFooterForSystem { get { return String.Format(@"<p class=MsoNormal>
<span style='font-size:12.0pt;font-family:""Arial"",""sans-serif"";mso-fareast-font-family:
""Times New Roman"";mso-fareast-theme-font:minor-fareast;mso-fareast-language:
RU;mso-no-proof:yes'>Если вы получили это письмо ошибочно или у вас возникли вопросы обращайтесь в
<a href=""mailto:Служба%20внедрения%20и%20сопровождения%20систем%20автоматизации"">Отдел автоматизации</a>
 по тел. 1180, 1200, 1204<o:p></o:p></p>
<p>или по E-mail: <a href=""mailto:lipaev@dinamika-avia.ru"">lipaev@dinamika-avia.ru</a></span></p>
<p><img src=""data:image/png;base64,{0}  alt=""АСУ ""Динамика""""></p>
			</span></body>", Logo_Base64); } }

        private string CreateMessageBody()
        {
            string bodyHeader = "";
            string bodyFooter = "";
            if (FromSystem)
            {
                bodyHeader = BodyHeaderForSystem;
                bodyFooter = BodyFooterForSystem;
            }
            return String.Format(
                @"<html>
{0}
  <p> {1} </p>
{2}
</html>", bodyHeader, BodyText, bodyFooter);
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

    /// <summary>
    /// Регистрационная карточка служебной записки в справочнике - Журнал регистрации служебных записок
    /// </summary>
    internal class ErrandRegistryJournalRecord
    {
        bool beginChangesApplied;
        bool hasModify;

        ReferenceObject _ReferenceObject;
        public ReferenceObject ReferenceObject
        {
            get
            {
                //if (_ReferenceObject.Changing) return _ReferenceObject.EditableObject;
                //else
                return _ReferenceObject;
            }
            private set
            {
                if (_ReferenceObject != value)
                    _ReferenceObject = value;
            }
        }


        public ErrandRegistryJournalRecord(ReferenceObject journalRecord_ro)
        {
            this.ReferenceObject = journalRecord_ro;
            beginChangesApplied = false;
            hasModify = false;
        }

        public ErrandRegistryJournalRecord(Guid guidRegistryObject)
        {
            //RegistryJournalRecordFinder.FindByGuid(guidRegistryObject);
            ReferenceObject = CreateNewRecord(guidRegistryObject.ToString());
            beginChangesApplied = false;
        }

        public void SetRegistryObject(IRegistryJournal objectForRregister)
        {

        }

        private ReferenceObject CreateNewRecord(string guidRegistryObjectString)
        {
            ReferenceObject newRecord = ErrandRegistryJournal_reference.CreateReferenceObject(ErrandRegistryCard_Class);
            newRecord[RJ_param_GuidRegistryObject_Guid].Value = guidRegistryObjectString;
            beginChangesApplied = true;
            hasModify = true;
            return newRecord;
        }

        public void SaveToDB()
        {
            if (hasModify) ReferenceObject.EndChanges();
        }

        internal void AddToExtraAccessForViewing(ReferenceObject user)
        {
            BeginChanges();
            ReferenceObject.AddLinkedObject(RJ_link_ExtraAccessForViewing_Guid, user);
            hasModify = true;
        }

        private void BeginChanges()
        {
            if (beginChangesApplied) return;
            else
            {
                ReferenceObject.BeginChanges(true);
                beginChangesApplied = true;
            }
        }

        private static Guid RegistryJournal_reference_Guid = new Guid("82f3b8b1-8880-4089-8370-80ca602c46c5");    //Guid справочника - "Журнал регистрации служебных записок"
        private static Guid RJ_ErrandRegistryCard_class_Guid = new Guid("f9f39854-69f1-402f-a13d-337e0482cd64");    //Guid типа "Карточка поручения" - справочника - "Журнал регистрации служебных записок"
        public static Guid RJ_link_Errand_Guid = new Guid("c2096bae-3a60-4269-b9f5-42e91ecc2ced");//связь 1:1 "Контроль поручений -> Журнал регистрации" со справочником "Служебные записки"
        public static Guid RJ_link_ExtraAccessForViewing_Guid = new Guid("bec2d0c6-9839-4fb7-ab44-ac664b48256e"); //связь N:N "Просмотр" со cправочником Группы и пользователи
        public static Guid RJ_link_Coordinators_UsersAndPositions_Guid = new Guid("dc34fc59-7358-4846-98e7-302c9d9dbbc8"); //связь N:N "Пользователи и должности -> Согласующие" со cправочником "Пользователи и должности"
        public static Guid RJ_param_GuidRegistryObject_Guid = new Guid("0479b7a3-b305-4bc4-a5e0-4c3250de1239"); //параметр "Guid зарегистрированного объекта" справочника - "Журнал регистрации служебных записок"


        #region Работа с БД

        //поля
        private Reference _ErrandRegistryJournalReference;
        private ReferenceInfo _ErrandRegistryJournal_ReferenceInfo;
        //Типы
        private static ClassObject _ErrandRegistryCard_Internal_class;
        private Reference GetReference(ref Reference reference, ReferenceInfo referenceInfo)
        {
            if (reference == null)
                reference = referenceInfo.CreateReference();


            return reference;
        }


        /// <summary>
        /// Справочник "Журнал регистрации служебных записок"
        /// </summary>
        private Reference ErrandRegistryJournal_reference
        {
            get
            {
                if (_ErrandRegistryJournalReference == null)

                    return GetReference(ref _ErrandRegistryJournalReference, ErrandRegistryJournal_ReferenceInfo);

                return _ErrandRegistryJournalReference;
            }
        }


        /// <summary>
        /// Информация о справочнике - Журнал регистрации служебных записок
        /// </summary>
        private ReferenceInfo ErrandRegistryJournal_ReferenceInfo
        {
            get { return GetReferenceInfo(ref _ErrandRegistryJournal_ReferenceInfo, RegistryJournal_reference_Guid); }
        }

        private ReferenceInfo GetReferenceInfo(ref ReferenceInfo referenceInfo, Guid referenceGuid)
        {
            if (referenceInfo == null)
                referenceInfo = Connection.ReferenceCatalog.Find(referenceGuid);

            return referenceInfo;
        }



        /// <summary>
        /// Тип - Карточка служебной записки
        /// </summary>
        private ClassObject ErrandRegistryCard_Class
        {
            get
            {
                if (_ErrandRegistryCard_Internal_class == null)
                    _ErrandRegistryCard_Internal_class = ErrandRegistryJournal_reference.Classes.Find(RJ_ErrandRegistryCard_class_Guid);

                return _ErrandRegistryCard_Internal_class;
            }
        }
        ServerConnection Connection { get { return ServerGateway.Connection; } }
        #endregion
    }

    static class RegistryJournalRecordFinder
    {
        static RegistryJournalRecordFinder()
        {
            // получение описания справочника
            ReferenceInfo referenceInfo = Connection.ReferenceCatalog.Find(new Guid("82f3b8b1-8880-4089-8370-80ca602c46c5"));
            // создание объекта для работы с данными
            Reference = referenceInfo.CreateReference();
        }
        public static ErrandRegistryJournalRecord FindById(int id)
        {
            ReferenceObject foundedObject = Reference.Find(id);
            if (foundedObject == null) return null;
            return new ErrandRegistryJournalRecord(foundedObject);
        }

        public static ErrandRegistryJournalRecord FindByGuid(string stringGuid)
        {
            return FindByGuid(new Guid(stringGuid));
        }

        public static ErrandRegistryJournalRecord FindByGuid(Guid guid)
        {
            ReferenceObject foundedObject = Reference.Find(Reference.ParameterGroup.OneToOneParameters.Find(ErrandRegistryJournalRecord.RJ_param_GuidRegistryObject_Guid), guid.ToString()).FirstOrDefault();
            if (foundedObject == null) return null;
            return new ErrandRegistryJournalRecord(foundedObject);
        }

        static Reference Reference;
        static ServerConnection Connection { get { return ServerGateway.Connection; } }
    }


}
