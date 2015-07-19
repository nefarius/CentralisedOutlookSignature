namespace CentralisedOutlookSignature.Service
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
            this.serviceProcessInstallerCOS = new System.ServiceProcess.ServiceProcessInstaller();
            this.serviceInstallerCOS = new System.ServiceProcess.ServiceInstaller();
            // 
            // serviceProcessInstallerCOS
            // 
            this.serviceProcessInstallerCOS.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
            this.serviceProcessInstallerCOS.Password = null;
            this.serviceProcessInstallerCOS.Username = null;
            // 
            // serviceInstallerCOS
            // 
            this.serviceInstallerCOS.ServiceName = "Centralised Outlook Signature";
            this.serviceInstallerCOS.StartType = System.ServiceProcess.ServiceStartMode.Automatic;
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.serviceProcessInstallerCOS,
            this.serviceInstallerCOS});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller serviceProcessInstallerCOS;
        private System.ServiceProcess.ServiceInstaller serviceInstallerCOS;
    }
}