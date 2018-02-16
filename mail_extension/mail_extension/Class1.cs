using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Mail;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Users;

namespace mail_extension
{
    /*
    class Program
    {
        void exemple()
        {
            List<ReferenceObject> mailRecipients = new List<ReferenceObject> { ServerGateway.Connection.References.Users.FindUser("Липаев Алексей Александрович") };

            HTML_MailContent MailContent = new HTML_MailContent("Просроченные работы");
            MailContent.BodyAddStrong("Вот мой телефон ");
            MailContent.BodyAddPhoneLink("89100151944");
            MailContent.FooterAdd("Если вы получили это письмо ошибочно или у вас возникли вопросы обращайтесь в по тел. 1180, 1200, 1204, 1206 или по E-mail ");
            MailContent.FooterAddEmailLink("kachkov@dinamika-avia.ru");
            HTML_MailMessage message = new HTML_MailMessage(MailContent, mailRecipients);
            message.Send();

        }
    }
    */

    /// <summary>
    /// HTML письмо
    /// ver 1.8
    /// </summary>
    public class HTML_MailMessage
    {
        public HTML_MailMessage()
        {
            Subject = string.Empty;
            BodyText = string.Empty;
            MailAttachments = new List<ReferenceObject>();
            MailRecipients = new List<ReferenceObject>();
            MailCopyRecipients = new List<ReferenceObject>();
        }

        public HTML_MailMessage(string messageSubject, string messageText, List<ReferenceObject> mailRecipients, List<ReferenceObject> works = null)
        {
            this.Subject = messageSubject;
            this.BodyText = messageText;
            this.MailRecipients = mailRecipients;

            if (MailAttachments == null)
                MailAttachments = new List<ReferenceObject>();

            if (works != null)
                this.MailAttachments = works;

            this.MailCopyRecipients = new List<ReferenceObject>();
        }

        public HTML_MailMessage(string messageSubject, string messageText, ReferenceObject mailRecipient, List<ReferenceObject> works = null)
                : this(messageSubject, messageText, new List<ReferenceObject>() { mailRecipient }, works)
        { }

        public HTML_MailMessage(HTML_MailContent HTML_emailDocument, List<ReferenceObject> mailRecipients, List<ReferenceObject> works = null)
           : this(HTML_emailDocument.Title, HTML_emailDocument.ToString(), mailRecipients, works)
        { }

        public HTML_MailMessage(HTML_MailContent HTML_emailDocument, ReferenceObject mailRecipient, List<ReferenceObject> works = null)
          : this(HTML_emailDocument.Title, HTML_emailDocument.ToString(), new List<ReferenceObject>() { mailRecipient }, works)
        { }

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
        //List<ReferenceObject> _MailCopyRecipients;
        public List<ReferenceObject> MailCopyRecipients { get; set; }
        public void Send()
        {
            MailMessage = new MailMessage(Connection.Mail.DOCsAccount);
            MailMessage.Subject = CleanString(Subject);
            MailMessage.SetBody(BodyText, MailBodyType.Html);
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
                //гуид абстрактного класса Файл справочника Файлы
                Guid fileClassGuid = new Guid("4731e1b6-b27e-4895-be2f-b8140316bfc0");
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

        static public bool FromSystem { get { return Connection.ClientView.GetUser().IsSystem; } }
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

    public class HTML_MailContent
    {
        // Поля
        private Header_HTML head;
        private Body_HTML body;
        private Footer_HTML footer;
        public string Title { get; set; }

        private string emailLink_HTMLtemplay = @"<a href =""mailto:{0}"" target=""_blank"" style=""{1}"">{0}</a>";
        private string phoneLink_HTMLtemplay = @"<a href =""tel:{0}"" target=""_blank"" style=""{1}"">{0}</a>";


        private string content_HTML;


        public HTML_MailContent(string title)
        {
            this.Title = title;

            InitializeDocument(title);
        }
        public HTML_MailContent(string title, string pathIconModule, string nameModule, string linkModule = null)
        {
            this.Title = title;
            InitializeDocument(title, pathIconModule, nameModule, linkModule);
        }
        private void InitializeDocument()
        {
            this.body = new Body_HTML();
            this.footer = new Footer_HTML();
        }

        private void InitializeDocument(string title)
        {
            this.Title = title;
            InitializeDocument();
            this.head = new Header_HTML(title);
        }
        private void InitializeDocument(string title, string pathIconModule, string nameModule, string linkModule = null)
        {
            this.Title = title;
            InitializeDocument();
            this.head = new Header_HTML(title, pathIconModule, nameModule, linkModule);
        }



        public void BodyAddEmailLink(string email, string Style = null)
        {
            var style = string.Empty;
            if (!string.IsNullOrEmpty(Style))
                style = Style;

            body.Content_HTML = string.Format(emailLink_HTMLtemplay, email, style);
        }

        public void BodyAddEmailLinkLine(string email, string style)
        {
            BodyAddEmailLink(email + "<br><br>", style);
        }

        public void BodyAddPhoneLink(string phone, string Style = null)
        {
            var style = string.Empty;
            if (!string.IsNullOrEmpty(Style))
                style = Style;

            body.Content_HTML = string.Format(phoneLink_HTMLtemplay, phone, style);
        }

        public void BodyAddPhoneLinkLine(string phone, string style)
        {
            FooterAddPhoneLink(phone + "<br><br>", style);
        }

        public void BodyAdd(string data)
        {
            body.Write(data);
        }

        public void BodyAddLine(string data)
        {
            body.WriteLine(data);
        }

        public void BodyAddStrong(string data)
        {
            BodyAdd("<strong>" + data + "</strong>");
        }
        public void BodyAddLineStrong(string data)
        {
            BodyAddStrong(data + "<br><br>");
        }

        public void BodyAdd(StringBuilder data)
        {
            BodyAdd(data.ToString());
        }

        public void BodyAddLine(StringBuilder data)
        {
            BodyAddLine(data.ToString());
        }

        public void BodyAddStrong(StringBuilder data)
        {
            BodyAdd("<strong>" + data.ToString() + "</strong>");
        }
        public void BodyAddLineStrong(StringBuilder data)
        {
            BodyAddStrong(data.ToString() + "<br><br>");
        }
        public void FooterAdd(string data)
        {
            footer.Write(data);
        }

        public void FooterAddLine(string data)
        {
            FooterAdd(data + "<br><br>");
        }
        public void FooterAddEmailLink(string email, string Style = null)
        {
            var style = string.Empty;
            if (!string.IsNullOrEmpty(Style))
                style = Style;

            footer.Content_HTML = string.Format(emailLink_HTMLtemplay, email, style);
        }

        public void FooterAddEmailLinkLine(string email, string style)
        {
            FooterAddEmailLink(email + "<br><br>", style);
        }
        public void FooterAddPhoneLink(string phone, string Style = null)
        {
            var style = string.Empty;
            if (!string.IsNullOrEmpty(Style))
                style = Style;

            footer.Content_HTML = string.Format(phoneLink_HTMLtemplay, phone, style);
        }

        public void FooterAddPhoneLinkLine(string phone, string style)
        {
            FooterAddPhoneLink(phone + "<br><br>", style);
        }


        private string Content_HTML
        {
            get
            {
                if (string.IsNullOrEmpty(content_HTML))
                    content_HTML = Head + Body + Footer;

                return content_HTML;
            }
        }

        public string Head
        {
            get
            {
                return head.Content_HTML.ToString();
            }
        }

        public string Body
        {

            get
            {
                return body.Content_HTML.ToString();
            }
            set
            {
                body.Content_HTML = value;
            }
        }

        public string Footer
        {
            get
            {
                return footer.Content_HTML.ToString();
            }
        }
        public override string ToString()
        {
            return Content_HTML.ToString();
        }

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
                                               <td> {3}&nbsp;</td>
                                               <td> {4}&nbsp;</td>
                                   </tr>", index++, sign.SignatureObjectType.Name, sign.UserName, signDate, sign.Resolution));
                }
                else continue;
            }
            table.Append("</table>");
            return table.ToString();
        }

        public static string GetLinkFor(ReferenceObject object_ro)
        {
            if (object_ro == null) return null;
            return String.Format(@"docs://{0}/OpenReferenceWindow/?refId={1}&objId={2}",
                                 ServerGateway.Connection.ConnectionParameters.GetServerAddress(), object_ro.Reference.Id, object_ro.SystemFields.Id);
        }
        public static string GetLinkFor(ReferenceObject object_ro, string nameLink, string StyleLink = null)
        {
            string styleLink = string.Empty;

            if (!string.IsNullOrEmpty(StyleLink))
                styleLink = StyleLink;

            string link = GetLinkFor(object_ro);

            return string.Format(@"<a href =""{0}"" target=""_blank"" style=""{2}"">{1}</a>", link, nameLink, styleLink);
        }

    }

    interface ICommonHTMLComponents
    {
        string Content_HTML { get; set; }
        void Write(string data);
        void WriteLine(string data);
    }

    class Header_HTML
    {
        StringBuilder content = new StringBuilder();

        public string Content_HTML
        {
            get
            {
                return content.ToString();
            }
            set
            {
                content.Append(value);
            }
        }

        StringBuilder templay = new StringBuilder().Append(@"
<!DOCTYPE HTML PUBLIC "" -//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd"">
<html>
   <head>
      <meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">
      <title>-</title>
   </head>
   <body style=""padding:0px;margin:0px;"" >
      <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" >
        {0}
      <!--header -->
      <tr>
         <td align=""center"" bgcolor=""#184263"" style=""border-bottom-width:1px;border-bottom-style:solid;border-bottom-color:#282f37;"">
               <table width=""80%"" border=""0"" cellspacing=""0"" cellpadding=""0"" height=""70"">
                  <tr>
                     <!-- заголовок --> 
                     <td align=""left"">
                        <div>
                           <font face=""Tahoma, Arial, Helvetica, sans-serif"" size=""3"" color=""#fffff"" style=""font-size:16px;"">
                           <span style=""font-family: Tahoma, Arial, Helvetica, sans-serif; font-size: 150%; color:#fffff; letter-spacing: 2px;"">
                           <strong>{1}</strong>
                           </span></font>
                        </div>
                     </td>
                     <!--заголовок END-->
                    {2}
                  </tr>
               </table>
               </td>
               </tr>
               <!--header END-->");

        string templaySystemText = @"<tr><!--preheader -->
          <td align=""center"" style=""background-color: #777f8c;  background-image: url(http://www.dinamika-avia.ru/bitrix/templates/defence_style_header_2017/images/top-panel-sprite_1.png)"">
            <table width=""80%"" border=""0"" cellspacing=""0"" cellpadding=""0"">
               <tr>
                  <td align=""left"" valign=""middle"" style=""line-height:20px;"">
                     <font face=""Tahoma, Arial, Helvetica, sans-serif"" size=""2"" color=""#ffffff"" style=""font-size:13px;"">
                     <span style=""font-family: Tahoma, Arial, Helvetica, sans-serif; font-size: 13px; color:#ffffff;line-height:20px;"">
                     Пожалуйста, не отвечайте на это письмо, так как оно сгенерировано автоматически.
                     </span></font>
                  </td>
               </tr>
            </table>
         </td>
      </tr><!--preheader END-->";


        string templayIcomModule = @"
                    <td align=""right""><!-- иконка модуля -->
                        <a  href=""{2}"" target=""_blank"" style=""color: #596167; font-family: Arial, Helvetica, sans-serif; font-size: 18px;""  >
                        <font face=""Arial, Helvetica, sans-serif; font-size: 18px;"" size=""4"" color=""#596167"">
                        <img src=""{0}""  height=""60"" alt=""{1}"" border=""0"" style=""display: block;"" /></font></a>
                     </td><!-- иконка модуля  end -->";

        public Header_HTML(string title) : this(title, string.Empty, string.Empty, string.Empty)
        {
        }

        public Header_HTML(string title, string pathIconModule, string nameModule, string LinkModule = null)
        {
            string logoModule = string.Empty;
            string linkModule = string.Empty;

            if (!string.IsNullOrEmpty(LinkModule))
                linkModule = LinkModule;

            if (!string.IsNullOrEmpty(pathIconModule))
                logoModule = string.Format(templayIcomModule, pathIconModule, nameModule, LinkModule);

            if (HTML_MailMessage.FromSystem)
            {
                content.AppendFormat(templay.ToString(), templaySystemText, title, logoModule, linkModule);
            }
            else
            {
                content.AppendFormat(templay.ToString(), string.Empty, title, logoModule, linkModule);
            }
        }



    }

    class Body_HTML : ICommonHTMLComponents
    {
        public string Content_HTML
        {
            get
            {
                return string.Format(templay.ToString(), content.ToString());
            }
            set
            {
                content.Append(value);
            }
        }


        StringBuilder templay = new StringBuilder().Append(@"
               <tr><!--content-->
                  <td align=""center"" bgcolor=""#ffffff"" style=""border-top-width:1px;border-top-style:solid;border-top-color:#ffffff;"">
                     <table width=""70%"" border=""0"" cellspacing=""0"" cellpadding=""0"">
                        <tr>
                           <td align=""left"">
                              <!--padding-->
                              <div style=""height: 35px; line-height:45px; font-size:40px;"">&nbsp;</div>
                                <div style=""line-height:24px; overflow-x:auto; ""><!-- datas-->
                             <font face=""Tahoma, Arial, Helvetica, sans-serif"" size=""3"" color=""#282f37"" style=""font-size:16px;"">
                               <div style=""line-height:24px;"">
                                    <span style=""font-family: Tahoma, Arial, Helvetica, sans-serif; font-size: 16px; color:#282f37;"">
                                    {0}
                                    </span>
                              </div>
                            </font>
                                </div><!-- datas END -->
                           </td>
                        </tr>
                     </table>
                     <!--padding-->
                     <div style=""height: 30px; line-height:30px; font-size:28px;""> &nbsp;</div>
                  </td>
               </tr><!--content END-->");

        StringBuilder content = new StringBuilder();

        public void Write(string data)
        {
            content.Append(data);
        }
        public void WriteLine(string data)
        {
            Write(data + "<br><br>");
        }

    }

    class Footer_HTML : ICommonHTMLComponents
    {

        public string Content_HTML
        {
            get
            {
                return string.Format(templay.ToString(), content.ToString());
            }
            set
            {
                content.Append(value);
            }
        }


        StringBuilder templay = new StringBuilder().Append(@"
               
               <tr> <!--footer-->
                  <td align=""center"" bgcolor=""#e7e9eb"">
                     <!--padding-->
                     <div style=""height: 30px; line-height:30px; font-size:28px;""> &nbsp;</div>
                     <table width=""80%"" border=""0"" cellspacing=""0"" cellpadding=""0"">
                        <tr>
                           <td  align=""left"" valign=""top"" style=""line-height:15px;"">
                              <font face=""Tahoma, Arial, Helvetica, sans-serif"" size=""2"" color=""#929ca8"" style=""font-size:13px;"">
                                <span style=""font-family: Tahoma, Arial, Helvetica, sans-serif; font-size: 13px; color:#929ca8;line-height:15px;"">
                                {0}
                                </span>
                              </font>
                              <!--logo-->
                              <table border=""0"" cellspacing=""0"" cellpadding=""0"">
                                 <a href=""http://www.dinamika-avia.ru/"" target=""_blank"" style=""color: #596167; font-family: Arial, Helvetica, sans-serif; font-size: 18px;"">
                                 <font face=""Arial, Helvetica, sans-serif; font-size: 18px;"" size=""4"" color=""#596167"">
                                 <img src=""X:/Подразделения/!Обмен/Качков/dinamika_tech_logo_blue_orange_ru_png.png"" width=""210"" height=""45"" alt=""Перейти на сайт «ЦНТУ &laquo;Динамика&raquo; :: Авиационные технологии ::»"" border=""0"" style=""display: block;""></font></a>
                              </table>
                              <!--logo-->
                           </td>
                           <td align=""center"">
                           </td>
                        </tr>
                     </table>
                     <!--padding-->
                     <div style=""height: 30px; line-height:30px; font-size:28px;""> &nbsp;</div>
                  </td>
               </tr><!--footer END-->
            </table>
   </body>
</html>");

        StringBuilder content = new StringBuilder();

        public void Write(string data)
        {
            content.Append(data);
        }
        public void WriteLine(string data)
        {
            Write(data + "<br><br>");
        }
    }
}
