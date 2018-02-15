using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Mail;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Users;

namespace TFLEXDocsClasses
{
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
}
