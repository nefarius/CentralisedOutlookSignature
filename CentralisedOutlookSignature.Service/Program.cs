using System;
using System.Configuration.Install;
using System.Reflection;
using System.ServiceProcess;

namespace CentralisedOutlookSignature.Service
{
    internal static class Program
    {
        /// <summary>
        ///     The main entry point for the application.
        /// </summary>
        private static void Main(string[] args)
        {
            if (Environment.UserInteractive)
            {
                try
                {
                    var parameter = string.Concat(args);
                    switch (parameter)
                    {
                        case "--install":
                            ManagedInstallerClass.InstallHelper(new[] {Assembly.GetExecutingAssembly().Location});
                            break;
                        case "--uninstall":
                            ManagedInstallerClass.InstallHelper(new[] {"/u", Assembly.GetExecutingAssembly().Location});
                            break;
                    }
                }
                catch
                {
                }
            }
            else
            {
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[]
                {
                    new CentralisedOutlookSignatureService()
                };
                ServiceBase.Run(ServicesToRun);
            }
        }
    }
}