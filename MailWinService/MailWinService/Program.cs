﻿using System.ServiceProcess;

namespace MailWinService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            var servicesToRun = new ServiceBase[] 
            { 
                new MailService() 
            };
            ServiceBase.Run(servicesToRun);
        }
    }
}
