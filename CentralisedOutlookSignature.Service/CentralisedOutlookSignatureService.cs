#region Usings
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Threading;
using CentralisedOutlookSignature.RxDefinitions;
using CentralisedOutlookSignature.Service.Properties;
using ReactiveProtobuf.Protocol;
using ReactiveSockets;
#endregion

namespace CentralisedOutlookSignature.Service
{
    public partial class CentralisedOutlookSignatureService : ServiceBase
    {
        #region Static variables
        private static readonly ReactiveListener RxServer = new ReactiveListener(Settings.Default.RxListenPort);
        private static readonly Dictionary<int, ProtobufChannel<RxMessage>> ClientList = new Dictionary<int, ProtobufChannel<RxMessage>>();
        private static readonly string WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory;
        private static FileSystemWatcher _fsWatch;
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private static readonly object SyncLock = new object();
        #endregion

        #region Ctor

        public CentralisedOutlookSignatureService()
        {
            InitializeComponent();
        }

        #endregion

        #region Service events

        protected override void OnStart(string[] args)
        {
            AppDomain.CurrentDomain.UnhandledException += CurrentDomainOnUnhandledException;

            Log.InfoFormat("Setting current directory to {0}", WorkingDirectory);
            Directory.SetCurrentDirectory(WorkingDirectory);

            if (!Directory.Exists(Settings.Default.SignatureRepository))
            {
                Log.ErrorFormat("Directory {0} does not exist!", Settings.Default.SignatureRepository);
                Stop();
                return;
            }

            _fsWatch = new FileSystemWatcher(Settings.Default.SignatureRepository, Settings.Default.WatcherFilter)
            {
                EnableRaisingEvents = true
            };

            _fsWatch.Changed += fileSystemWatcherCOS_Changed;

            RxServer.Connections.Subscribe(socket =>
            {
                Log.InfoFormat("New client {0} connected", socket.GetHashCode());

                socket.Disconnected += (sender, eventArgs) =>
                {
                    Log.InfoFormat("Client {0} disconnected", socket.GetHashCode());

                    lock (SyncLock)
                    {
                        var client = ClientList.First(c => c.Key == socket.GetHashCode());
                        ClientList.Remove(client.Key);
                    }
                };

                lock (SyncLock)
                {
                    ClientList.Add(socket.GetHashCode(), new ProtobufChannel<RxMessage>(socket));
                }
            });

            RxServer.Start();
        }

        protected override void OnStop()
        {
            Log.InfoFormat("Shutting down service, closing remaining connections...");
        }

        #endregion

        #region Event handlers

        private static void CurrentDomainOnUnhandledException(object sender, UnhandledExceptionEventArgs unhandledExceptionEventArgs)
        {
            // shouldn't happen :3
            var ex = (Exception)unhandledExceptionEventArgs.ExceptionObject;

            Log.ErrorFormat("Unhandled exception: {0}", ex);
        }
        
        private static void fileSystemWatcherCOS_Changed(object sender, FileSystemEventArgs e)
        {
            Log.InfoFormat("File {0} has been changed", e.FullPath);

            Log.Info("Waiting 3 seconds bevor pushing out message...");
            Thread.Sleep(3000);

            lock (SyncLock)
            {
                foreach (var protobufChannel in ClientList)
                {
                    protobufChannel.Value.SendAsync(new RxMessage());
                }
            }
        }

        #endregion
    }
}