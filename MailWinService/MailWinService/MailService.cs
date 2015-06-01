using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Configuration;
using System.Net.Mail;
using System.ServiceProcess;
using System.Threading;

namespace MailWinService
{
    public partial class MailService : ServiceBase
    {
        private EventLog _eventLog;
        private Timer _timer;
        public MailService()
        {
            InitializeComponent();
        }

        public void MyServiceOnStart()
        {
            //MailSync(null);
            //Debug işlemleri burada
        }

        protected override void OnStart(string[] args)
        {

            // Timer çalışmadan önce 1(1000ms) saniye bekelr
            // Her 1(60000ms) dakikada tetiklenir
            // MailSync metotu null parametresi ile çalıştırılır
            _timer = new Timer(MailSync, null, 1000, 60000);

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
        protected void MailSync(object sender)
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

                    DbOperation.Execute("update MailBox set IsSent='true' where Id="+mailBox["Id"].ToString());
                }
            }
        }
    }
}
