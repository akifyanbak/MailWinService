using System;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
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

        protected void MailSync(object sender)
        {
            //Mail işlemleri
        }
    }
}
