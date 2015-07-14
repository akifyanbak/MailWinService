using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using ljk.data.Context;
using ljk.data.Domain;
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
            //InboxMailSync(null);
            OutboxMailSync(null);
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
            try
            {
                var ljkContext = new LjkContext();
                var mailBoxes = ljkContext.MailBoxes.Where(m => !m.IsSent && !m.IsInbox).ToList();
                var mailAccounts = ljkContext.Mails.ToList();

                foreach (var mailBox in mailBoxes)
                {
                    var mailAccount = mailAccounts.FirstOrDefault(m => m.MailAddress == mailBox.Sender);
                    if (mailAccount != null)
                    {
                        var smtpClient = new SmtpClient(mailAccount.OutgoingMailServer, mailAccount.OutgoingMailPort)
                        {
                            Credentials = new NetworkCredential(mailAccount.MailAddress, mailAccount.Password),
                            EnableSsl = true
                        };
                        var mailMessage = new MailMessage { From = new MailAddress(mailAccount.MailAddress, mailAccount.Username) };
                        mailMessage.To.Add(mailBox.MailTo);
                        mailMessage.Subject = mailBox.Subject;
                        mailMessage.IsBodyHtml = true;
                        mailMessage.Body = mailBox.Content;
                        foreach (var attach in mailBox.Attaches)
                        {
                            Attachment data = new Attachment(attach.FilePath);
                            mailMessage.Attachments.Add(data);
                        }
                        smtpClient.Send(mailMessage);
                        smtpClient.Dispose();
                        mailBox.IsSent = true;
                        ljkContext.SaveChanges();
                    }
                }
            }
            catch (Exception e)
            {
                Logging(e.Message);
            }

        }

        protected void InboxMailSync(object sender)
        {
            try
            {
                var ljkContext = new LjkContext();
                var mailAccounts = ljkContext.Mails.ToList();

                foreach (var mailAccount in mailAccounts)
                {
                    var mails = FetchAllMessages(
                        mailAccount.IncomingMailServer,
                      mailAccount.IncomingMailPort,
                       true,
                       mailAccount.MailAddress,
                       mailAccount.Password);
                    foreach (var message in mails)
                    {
                        if (message != null)
                        {
                            // Mail veritabanında kayıtlı değilse girecek
                            if (!ljkContext.MailBoxes.Any(m => m.EmailId == message.Headers.MessageId))
                            {
                                var attaches = new List<Attach>();
                                var mailbox = new MailBox()
                                {
                                    Content = MessageText(message.MessagePart),
                                    EmailId = message.Headers.MessageId,
                                    Sender = message.Headers.From.Address,
                                    MailTo = mailAccount.MailAddress,
                                    IsInbox = true,
                                    Subject = message.Headers.Subject
                                };

                                if (message.MessagePart.IsMultiPart)
                                {
                                    foreach (var part in message.MessagePart.MessageParts)
                                    {
                                        if (part.IsAttachment)
                                        {
                                            var attach = new Attach
                                            {
                                                FileName = Guid.NewGuid() + "-" + part.FileName,
                                                FileData = part.Body,
                                                MediaType = part.ContentType.MediaType,
                                                FilePath = @"D:\Uploads\Ljk\Mail\"
                                            };
                                            SaveData(attach.FilePath, attach.FileName, part.Body);
                                            attaches.Add(attach);
                                        }
                                    }
                                }
                                mailbox.Attaches = attaches;
                                ljkContext.MailBoxes.Add(mailbox);
                                ljkContext.SaveChanges();
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Logging(e.Message);
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

        private void Logging(string logText)
        {
            const string path = "C:\\MailService\\Logging.txt";
            using (StreamWriter sw = (File.Exists(path)) ? File.AppendText(path) : File.CreateText(path))
            {
                sw.WriteLine(DateTime.Now + " - Hata: " + logText);
            }
        }

        protected bool SaveData(string filepath, string fileName, byte[] data)
        {
            try
            {
                // Create a new stream to write to the file
                if (!Directory.Exists(filepath)) Directory.CreateDirectory(filepath);
                var writer = new BinaryWriter(File.OpenWrite(filepath + fileName));

                // Writer raw data                
                writer.Write(data);
                writer.Flush();
                writer.Close();
            }
            catch
            {
                Logging("Dosyalar yüklenemedi");
                return false;
            }

            return true;
        }

        public string MessageText(MessagePart messagePart)
        {
            if (messagePart.IsText)
            {
                return messagePart.GetBodyAsText();
            }
            foreach (var part in messagePart.MessageParts)
            {
                return part.IsText ? part.GetBodyAsText() : MessageText(part);
            }
            return "";
        }
    }
}
