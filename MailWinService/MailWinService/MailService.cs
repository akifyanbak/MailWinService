using System.Diagnostics;
using System.ServiceProcess;

namespace MailWinService
{
    public partial class MailService : ServiceBase
    {
        private EventLog _eventLog;

        public MailService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
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
    }
}
