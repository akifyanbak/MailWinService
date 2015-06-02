namespace MailWinService
{
    partial class ProjectInstaller
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.mailSyncServiceProcessInstaller = new System.ServiceProcess.ServiceProcessInstaller();
            this.mailSyncInstallerService = new System.ServiceProcess.ServiceInstaller();
            // 
            // mailSyncServiceProcessInstaller
            // 
            this.mailSyncServiceProcessInstaller.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
            this.mailSyncServiceProcessInstaller.Password = null;
            this.mailSyncServiceProcessInstaller.Username = null;
            // 
            // mailSyncInstallerService
            // 
            this.mailSyncInstallerService.ServiceName = "MailService";
            this.mailSyncInstallerService.StartType = System.ServiceProcess.ServiceStartMode.Automatic;
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.mailSyncServiceProcessInstaller,
            this.mailSyncInstallerService});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller mailSyncServiceProcessInstaller;
        private System.ServiceProcess.ServiceInstaller mailSyncInstallerService;
    }
}