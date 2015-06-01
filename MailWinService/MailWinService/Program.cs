using System.ServiceProcess;

namespace MailWinService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            //Debug yapmak için false yapılması gerekiyor
            const bool debug = true;

            if (debug)
            {
                var servicesToRun = new ServiceBase[] 
                { 
                    new MailService() 
                };
                ServiceBase.Run(servicesToRun);
            }
            else
            {
                MailService myService = new MailService();
                myService.MyServiceOnStart();
            }
        }
    }
}
