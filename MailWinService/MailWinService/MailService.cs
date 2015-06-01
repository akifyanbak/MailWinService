using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using OpenPop.Mime;
using OpenPop.Pop3;

namespace MailWinService
{
    public partial class MailService : ServiceBase
    {
        private EventLog _eventLog;
        private Timer _timerOutbox;
        private Timer _timerInbox;
        public MailService()
        {
            InitializeComponent();
        }

        public void MyServiceOnStart()
        {
            InboxMailSync(null);
            //MailSync(null);
            //Debug işlemleri burada
        }

        protected override void OnStart(string[] args)
        {

            // Timer çalışmadan önce 1(1000ms) saniye bekelr
            // Her 1(60000ms) dakikada tetiklenir
            // MailSync metotu null parametresi ile çalıştırılır
            _timerOutbox = new Timer(OutboxMailSync, null, 1000, 60000);

            _timerInbox = new Timer(InboxMailSync, null, 1000, 60000);

            _eventLog = new EventLog();
            if (!EventLog.SourceExists("SampleSource"))
            {
                /* Ilk parametre ile, "Log Test" ismi altinda tutulacak 
                 * Log bilgilerinin kaynak ismi belirleniyor. Daha sonra 
                 * bu kaynak ismi _eventLog isimli nesnemizin Source özelligine ataniyor.*/
                EventLog.CreateEventSource("SampleSource", "Log");
            }
            _eventLog.Source = "SampleSource";

            /* Log olarak ilk parametrede belirtilen mesaj yazilir. 
             * Log'un tipi ise ikinci parametrede görüldügü gibi Information'dir.*/
            _eventLog.WriteEntry("Our service was started.", EventLogEntryType.Information);
        }

        protected override void OnStop()
        {
            _eventLog = new EventLog();
            if (!EventLog.SourceExists("SampleSource"))
            {
                EventLog.CreateEventSource("SampleSource", "Log");
            }
            _eventLog.Source = "SampleSource";
            _eventLog.WriteEntry("Our service was stopped.");

        }

        /// <summary>
        /// Mail senkron işlemi. sender dışardan bişey lazım olursa diye ekledim :)
        /// </summary>
        /// <param name="sender"></param>
        protected void OutboxMailSync(object sender)
        {
            var mailBoxes = DbOperation.Data("Select * from MailBox where IsSent='false' and IsDeleted='false' and IsInbox='false' and Sender in (select Mail.MailAddress from Mail)");
            var mailAccounts = DbOperation.Data("Select * from Mail");

            foreach (DataRow mailBox in mailBoxes.Rows)
            {
                var mailAccount = (from DataRow a in mailAccounts.Rows
                                   where mailBox != null && a["MailAddress"].ToString() == mailBox["Sender"].ToString()
                                   select a).FirstOrDefault();
                if (mailAccount != null)
                {
                    var smtpClient = new SmtpClient(mailAccount["OutGoingMailServer"].ToString(), Convert.ToInt32(mailAccount["OutGoingMailPort"]))
                    {
                        Credentials = new NetworkCredential(mailAccount["MailAddress"].ToString(), mailAccount["Password"].ToString()),
                        EnableSsl = true
                    };
                    var mailMessage = new MailMessage { From = new MailAddress(mailAccount["MailAddress"].ToString(), mailAccount["Username"].ToString()) };
                    mailMessage.To.Add(mailBox["MailTo"].ToString());
                    mailMessage.Subject = mailBox["Subject"].ToString();
                    mailMessage.IsBodyHtml = true;
                    mailMessage.Body = mailBox["Content"].ToString();
                    smtpClient.Send(mailMessage);
                    smtpClient.Dispose();
                    DbOperation.Execute("update MailBox set IsSent='true' where Id=" + mailBox["Id"].ToString());

                }
            }

        }

        protected void InboxMailSync(object sender)
        {
            var mailAccounts = DbOperation.Data("Select * from Mail");
            foreach (DataRow mailAccount in mailAccounts.Rows)
            {
                string mailAddress = mailAccount["MailAddress"].ToString();
                var mails = FetchAllMessages(
                    mailAccount["IncomingMailServer"].ToString(),
                   Convert.ToInt32(mailAccount["IncomingMailPort"]),
                   true,
                   mailAddress,
                   mailAccount["Password"].ToString());
                foreach (var message in mails)
                {
                    var DbMail = DbOperation.Data("select * from Mailbox where EmailId like '" + message.Headers.MessageId + "'");
                    if (DbMail.Rows.Count == 0)
                    {
                        string command = "insert into MailBox ([Subject],[Content],[MailTo],[IsSent],[Sender],[IsInbox],[CreatedDate],[IsActive],[IsDeleted],[EmailId])" +
                  "values (@Subject,@Content,@MailTo,@IsSent,@Sender,@IsInbox,GETDATE(),@IsActive,@IsDeleted,@EmailId)";
                        string[] parameters = { "@Subject", "@Content", "@MailTo", "@IsSent", "@Sender", "@IsInbox", "@IsActive", "@IsDeleted", "@EmailId" };
                        string[] values =
                    {
                        message.Headers.Subject,
                        message.MessagePart.IsText?message.MessagePart.GetBodyAsText():Encoding.UTF8.GetString(message.MessagePart.MessageParts.FirstOrDefault().Body),
                         mailAddress,
                         "false",
                         message.Headers.From.Address,
                         "false",
                         "true",
                         "false",
                         message.Headers.MessageId
                    };
                        DbOperation.Execute(parameters, values, command);
                    }
                }
            }
        }

        public List<Message> FetchAllMessages(string hostname, int port, bool useSsl, string username, string password)
        {
            // The client disconnects from the server when being disposed
            using (Pop3Client client = new Pop3Client())
            {
                // Connect to the server
                client.Connect(hostname, port, useSsl);

                // Authenticate ourselves towards the server
                client.Authenticate(username, password);

                // Get the number of messages in the inbox
                int messageCount = client.GetMessageCount();

                // We want to download all messages
                List<Message> allMessages = new List<Message>(messageCount);

                // Messages are numbered in the interval: [1, messageCount]
                // Ergo: message numbers are 1-based.
                // Most servers give the latest message the highest number
                for (int i = messageCount; i > 0; i--)
                {
                    allMessages.Add(client.GetMessage(i));
                }

                // Now return the fetched messages
                return allMessages;
            }
        }
    }
}
