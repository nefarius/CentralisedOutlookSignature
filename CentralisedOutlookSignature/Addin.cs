#region Usings
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Reactive.Concurrency;
using System.Reactive.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using CentralisedOutlookSignature.RxDefinitions;
using CSScriptLibrary;
using IniParser;
using IniParser.Model;
using log4net;
using log4net.Config;
using Libarius.Active_Directory;
using NetOffice.OfficeApi.Tools;
using NetOffice.Tools;
using NetOffice.WordApi.Enums;
using ReactiveProtobuf.Protocol;
using ReactiveSockets;
using Word = NetOffice.WordApi;
using Outlook = NetOffice.OutlookApi;
using Office = NetOffice.OfficeApi;
using VBIDE = NetOffice.VBIDEApi;
using Cfg = CentralisedOutlookSignature.Properties.Settings;
#endregion

namespace CentralisedOutlookSignature
{
    [COMAddin("Zentralisierte Signatur für Outlook", "Zentralisierte Signatur für Outlook", 3),
     ProgId("CentralisedOutlookSignature.Addin"), Guid("76CB6240-56A9-4E79-BC3F-D032DA7DBD17")]
    [RegistryLocation(RegistrySaveLocation.LocalMachine), CustomUI("CentralisedOutlookSignature.RibbonUI.xml")]
    [MultiRegister(RegisterIn.Outlook)]
    public class Addin : COMAddin
    {
        #region Variables

        private static string _dllFilePath;

        internal Office.IRibbonUI RibbonUI { get; private set; }

        #endregion

        #region Constructor

        public Addin()
        {
            // get absolute path to DLL config file
            var cfgFileName = string.Concat(Path.GetFileName(MyLocation), ".config");
            _dllFilePath = Path.GetDirectoryName(MyLocation) ?? string.Empty;
            var cfgFilePath = Path.Combine(_dllFilePath, cfgFileName);

            // fatal error if no configuration file
            if (!File.Exists(cfgFilePath))
            {
                MessageBox.Show(string.Format("Configuration file not found: {0}", cfgFilePath), "Fatal error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            /* reset current app domains configuration file to assemblys file
             * default is host EXE config file, so change is needed
             * */
            AppConfig.Change(cfgFilePath);
            // reload settings properties
            Cfg.Default.Reload();
            // initialize log4net subsystem
            XmlConfigurator.Configure(new FileInfo(cfgFilePath));
            // create logger with new settings
            _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

            /* locate GlobalSettings.ini
             * if path is absolute, use it
             * if path is relative, search in DLL directory
             * */
            var iniFilePath = Path.IsPathRooted(Cfg.Default.GlobalSettingsIniFilePath)
                ? Cfg.Default.GlobalSettingsIniFilePath
                : Path.Combine(_dllFilePath, Cfg.Default.GlobalSettingsIniFilePath);

            // abort and log error if file not found
            if (!File.Exists(iniFilePath))
            {
                _log.FatalFormat("Global configuration file {0} not found", iniFilePath);
                return;
            }

            // load global configuration
            var parser = new FileIniDataParser();
            _cfg = parser.ReadFile(iniFilePath);

            // initialize reactive client
            _rxClient = new ReactiveClient(_cfg["Rx"]["RxServerHost"], int.Parse(_cfg["Rx"]["RxServerPort"]));

            // abort if net45 is not found
            if (!IsNet45OrNewer())
            {
                _log.ErrorFormat(".NET 4.5 features not found on host {0}", Environment.MachineName);
                return;
            }

            // hook events
            OnStartupComplete += Addin_OnStartupComplete;
            OnDisconnection += Addin_OnDisconnection;
            OnConnection += Addin_OnConnection;
        }

        #endregion
        
        #region Helpers

        /// <summary>
        ///     Checks if .NET 4.5 features are available on this system.
        /// </summary>
        /// <returns>True if .NET 4.5 available, false otherwise.</returns>
        private static bool IsNet45OrNewer()
        {
            // Class "ReflectionContext" exists from .NET 4.5 onwards.
            return Type.GetType("System.Reflection.ReflectionContext", false) != null;
        }

        #endregion

        #region Main logic

        /// <summary>
        ///     Main logic to update signatures.
        /// </summary>
        /// <param name="force">If true, skip timestamp check. If false, skip update if template ain't new.</param>
        private async void UpdateSignatures(bool force = false)
        {
            try
            {
                var denied = Guid.Parse(_cfg["ActiveDirectory"]["SignatureUpdateDeniedGuid"]);

                if (denied != Guid.Empty &&
                    AdHelper.IsUserInGroup(denied))
                {
                    _log.InfoFormat("Current user {0} is not allowed to receive signatures", Environment.UserName);
                    return;
                }
            }
            catch (Exception adex)
            {
                _log.ErrorFormat("Failed to check group membership: {0}", adex);
                return;
            }

            #region Preparations

            // local relative office signature path
            const string relativeSigPath = @"Microsoft\Signatures";
            _log.DebugFormat("relativeSigPath = {0}", relativeSigPath);

            // local absolute path to signature folder
            var localSignaturePath = Path.Combine(AppDataRoaming, relativeSigPath);
            _log.DebugFormat("localSignaturePath = {0}", localSignaturePath);

            // a fresh user profile might not have the signature directory...
            if (!Directory.Exists(localSignaturePath))
            {
                _log.InfoFormat("Local signature directory {0} doesn't exist, creating...", localSignaturePath);

                try
                {
                    // ...so let's create it...
                    Directory.CreateDirectory(localSignaturePath);
                }
                catch (IOException ioex)
                {
                    // ...and log error if unsuccessful
                    _log.ErrorFormat("Couldn't create local signature directory: {0}", ioex);
                    return;
                }

                _log.InfoFormat("Local signature directory {0} created", localSignaturePath);
            }

            var signatureRepository = _cfg["Main"]["SignatureRepository"];

            // validate that signature repository folder exists...
            if (!Directory.Exists(signatureRepository))
            {
                _log.ErrorFormat("Signature repository ({0}) not found!", signatureRepository);
                // ...and abort if not
                return;
            }

            #endregion

            // loop through all connected mailbox accounts
            foreach (var mailAddress in _mailAddresses)
            {
                #region File-specific tasks

                // get currently logged on user
                var username = Environment.UserName;
                _log.DebugFormat("username = {0}", username);

                _log.DebugFormat("mailAddress = {0}", mailAddress);
                var signatures = Directory.GetFiles(signatureRepository, _cfg["Main"]["TemplateFilesFilter"]);

                /* Lookup signatures:
                 * 1. username.docx
                 * 2. user@domain.tld.docx
                 * 3. domain.tld.docx
                 * */
                var signatureSource = signatures.FirstOrDefault(s => Path.GetFileNameWithoutExtension(s)
                    .ToLower()
                    .Equals(username.ToLower())) ??
                                      signatures.FirstOrDefault(s => Path.GetFileNameWithoutExtension(s)
                                          .ToLower()
                                          .Equals(mailAddress.Address)) ??
                                      signatures.FirstOrDefault(s => Path.GetFileNameWithoutExtension(s)
                                          .ToLower()
                                          .Equals(mailAddress.Host));

                // log a warning if no matching template was found in repository
                if (string.IsNullOrEmpty(signatureSource))
                {
                    _log.WarnFormat("No signature template found for user {0} or mail address {1}, skipping...",
                        Environment.UserName, mailAddress.Address);
                    continue;
                }

                _log.DebugFormat("signatureSource = {0}", signatureSource);

                // append timestamp to file name to avoid overwriting locked files
                var sigName = string.Format("{0}_{1}{2}", Path.GetFileNameWithoutExtension(signatureSource),
                    DateTime.Now.ToString("yyyyMMddHHmmssffff"),
                    Path.GetExtension(signatureSource));
                _log.DebugFormat("sigName = {0}", sigName);

                var localSigName = Path.GetFileNameWithoutExtension(sigName);
                _log.DebugFormat("localSigName = {0}", localSigName);

                var remoteFileTimestamp = File.GetLastWriteTime(signatureSource);
                _log.DebugFormat("remoteFileTimestamp = {0}", remoteFileTimestamp);

                var localTimestampFile = Path.Combine(localSignaturePath,
                    string.Format("{0}_{1}", mailAddress.Address, remoteFileTimestamp.ToString("yyyyMMddHHmmssffff")));
                _log.DebugFormat("localTimestampFile = {0}", localTimestampFile);

                // if a timestamp file exists, skip processing the template
                if (!force && File.Exists(localTimestampFile))
                {
                    _log.InfoFormat("Signature is already up-to-date, skipping...");
                    return;
                }

                // copy template
                try
                {
                    _log.InfoFormat("Trying to copy file from \"{0}\" to \"{1}\"",
                        signatureSource, Path.Combine(localSignaturePath, sigName));
                    // copy template to local signature path
                    File.Copy(signatureSource, Path.Combine(localSignaturePath, sigName), true);
                    _log.Info("File copied successfully");
                }
                catch (IOException ioex)
                {
                    _log.ErrorFormat("Error copying signature: {0}", ioex);
                    return;
                }

                #endregion

                #region Active Directory

                // build LDAP filter
                var filter = string.Format(_cfg["ActiveDirectory"]["LdapFilter"], username);

                // create new directory searcher
                var searcher = new DirectorySearcher
                {
                    Filter = filter
                };

                DirectoryEntry adUser;

                try
                {
                    // query active directory async
                    adUser = await Task.Run(() => searcher.FindOne().GetDirectoryEntry());
                }
                catch (Exception ex)
                {
                    _log.ErrorFormat("Couldn't get directory entry: {0}", ex);
                    return;
                }

                #endregion

                #region MS Word actions

                var msWord = await Task.Run(() => new Word.Application());

                // absolute local path to template
                var fullTemplatePath = Path.Combine(localSignaturePath, sigName);

                try
                {
                    // open template
                    _log.DebugFormat("fullPath = {0}", fullTemplatePath);
                    await Task.Run(() => msWord.Documents.Open(fullTemplatePath));
                    // path to script
                    var tagDefinitionsFile = Path.Combine(_dllFilePath, @"Scripts\TagDefinitions.cs");
                    // check existance
                    if (!File.Exists(tagDefinitionsFile))
                    {
                        _log.ErrorFormat("TagDefinitions.cs not found");
                        return;
                    }

                    // load script content
                    var tagDefinitions = File.ReadAllText(tagDefinitionsFile);
                    // load assemblies
                    CSScript.Evaluator.ReferenceAssembliesFromCode(tagDefinitions);
                    // load dictionary with replacement variables from script
                    dynamic script = CSScript.Evaluator.LoadCode(tagDefinitions);
                    // get dictionary
                    var replaceDictionary = script.Initialize(adUser, _log);

                    // replace variables from template with AD content
                    foreach (var replace in replaceDictionary)
                    {
                        var findText = replace.Key;
                        _log.DebugFormat("findText = {0}", findText);

                        var replaceText = replace.Value;
                        _log.DebugFormat("replaceText = {0}", replaceText);

                        if (string.IsNullOrEmpty(replaceText))
                        {
                            if (msWord.Selection.Find.Execute(findText, false, true))
                                msWord.Selection.Range.Paragraphs[1].Range.Delete();
                            continue;
                        }

                        // check for hyperlinks and replace link target
                        var link = msWord.ActiveDocument.Hyperlinks.FirstOrDefault(h => h.TextToDisplay == findText);
                        if (link != null)
                        {
                            _log.DebugFormat("link = {0}", link);
                            link.Address = link.Address.Replace(_cfg["Templates"]["LinkHrefPlaceholder"], replaceText);
                        }

                        // find and replace placeholders
                        msWord.Selection.Find.Execute(findText, false, true, false, false, false, true, 1, false,
                            replaceText, 2);
                    }

                    _log.Info("Signature created from template");

                    // get begin and end of document content
                    object start = msWord.Application.ActiveDocument.Content.Start;
                    object end = msWord.Application.ActiveDocument.Content.End;
                    // select entire document
                    msWord.Application.ActiveDocument.Range(start, end).Select();

                    var oOpt = msWord.Application.EmailOptions;
                    var oSig = oOpt.EmailSignature;
                    var cSig = oSig.EmailSignatureEntries;
                    var oSel = msWord.Application.Selection;

                    // set default signature
                    var sigDisplayName = Path.GetFileNameWithoutExtension(signatureSource);
                    cSig.Add(sigDisplayName, oSel.Range);
                    oSig.NewMessageSignature = sigDisplayName;
                    oSig.ReplyMessageSignature = sigDisplayName;
                }
                catch (COMException comex)
                {
                    _log.ErrorFormat("COM-Interop failed: {0}", comex);
                    _log.ErrorFormat("\tInner exception: {0}", comex.InnerException);
                }
                catch (Exception ex)
                {
                    _log.ErrorFormat("Error creating signature: {0}", ex);
                }
                finally
                {
                    // dispose word process (don't save changes)
                    msWord.Quit(false);
                    msWord.Dispose();
                    msWord = null;

                    // TODO: add error handling!
                    // remove temporary template
                    File.Delete(fullTemplatePath);
                }

                #endregion

                // create timestamp file in AppData
                File.Create(localTimestampFile);
                _log.InfoFormat("Created timestamp file {0}", localTimestampFile);
            }

            MessageBox.Show(AddinResource.SignaturesUpdatedText, AddinResource.SignaturesUpdatedHeader,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        ///     Main logic to update template settings.
        /// </summary>
        private static async void UpdateTemplateSettings()
        {
            _log.InfoFormat("Updating template settings from {0}", Cfg.Default.GlobalSettingsIniFilePath);

            var msWord = await Task.Run(() => new Word.Application());

            try
            {
                var oOpt = msWord.Application.EmailOptions;

                // set letter templates
                oOpt.ReplyStyle.Font.Name = _cfg["Styles"]["ReplyStyleFontName"];
                oOpt.ReplyStyle.Font.Size = int.Parse(_cfg["Styles"]["ReplyStyleFontSize"]);
                oOpt.ReplyStyle.Font.Color =
                    (WdColor)Enum.Parse(typeof(WdColor), _cfg["Styles"]["ReplyStyleFontColor"]);
                oOpt.ReplyStyle.Font.Bold = int.Parse(_cfg["Styles"]["ReplyStyleFontBold"]);
                oOpt.ReplyStyle.Font.Italic = int.Parse(_cfg["Styles"]["ReplyStyleFontItalic"]);
                oOpt.ReplyStyle.Font.Underline =
                    (WdUnderline)Enum.Parse(typeof(WdUnderline), _cfg["Styles"]["ReplyStyleFontUnderline"]);

                oOpt.ComposeStyle.Font.Name = _cfg["Styles"]["ComposeStyleFontName"];
                oOpt.ComposeStyle.Font.Size = int.Parse(_cfg["Styles"]["ComposeStyleFontSize"]);
                oOpt.ComposeStyle.Font.Color =
                    (WdColor)Enum.Parse(typeof(WdColor), _cfg["Styles"]["ComposeStyleFontColor"]);
                oOpt.ComposeStyle.Font.Bold = int.Parse(_cfg["Styles"]["ComposeStyleFontBold"]);
                oOpt.ComposeStyle.Font.Italic = int.Parse(_cfg["Styles"]["ComposeStyleFontItalic"]);
                oOpt.ComposeStyle.Font.Underline =
                    (WdUnderline)Enum.Parse(typeof(WdUnderline), _cfg["Styles"]["ComposeStyleFontUnderline"]);
            }
            catch (COMException comex)
            {
                _log.ErrorFormat("COM-Interop failed: {0}", comex);
                _log.ErrorFormat("\tInner exception: {0}", comex.InnerException);
            }
            catch (Exception ex)
            {
                _log.ErrorFormat("Error creating signature: {0}", ex);
            }
            finally
            {
                // dispose word process (don't save changes)
                msWord.Quit(false);
                msWord.Dispose();
                msWord = null;
            }

            _log.Info("Finished template update");
        }

        #endregion

        #region Add-in events

        private void Addin_OnConnection(object application, ext_ConnectMode connectMode, object addInInst,
            ref Array custom)
        {
            _log.Info("Loading add-in");

            _myHost = Application as Outlook.Application;
            if (_myHost == null)
            {
                _log.ErrorFormat("Couldn't access host application as Microsoft Outlook");
                return;
            }

            _log.InfoFormat("Host version: {0}", _myHost.Version);

            _nsMapi = _myHost.GetNamespace("MAPI") as Outlook.NameSpace;
            if (_nsMapi == null)
            {
                _log.ErrorFormat("Couldn't get MAPI namespace");
                return;
            }

            // the following block works with Outlook 2010 or newer
            if (Version.Parse(_myHost.Version).Major > 12)
            {
                // fetch mail addresses currently configured
                _mailAddresses = (from store in _nsMapi.Stores
                                  from account in _myHost.Session.Accounts
                                  where account.DeliveryStore.StoreID == store.StoreID
                                  select new MailAddress(account.SmtpAddress.ToLower())).ToList();
            }
            else // this block is used on Outlook 2007
            {
                _mailAddresses = (from account in _myHost.Session.Accounts
                                  select new MailAddress(account.SmtpAddress.ToLower())).ToList();
            }

            // load template preferences at startup
            UpdateTemplateSettings();

            // check for signature updates
            UpdateSignatures();
        }

        private void Addin_OnStartupComplete(ref Array custom)
        {
            _log.InfoFormat("Host application ({0}) startup complete", (Application as Outlook.Application).Name);

            if (_rxClient == null)
            {
                _log.ErrorFormat("Rx connection invalid, can't listen for updates");
            }
            else
            {
                var channel = new ProtobufChannel<RxMessage>(_rxClient);

                channel.Receiver.SubscribeOn(TaskPoolScheduler.Default).Subscribe(message =>
                {
                    // wait a random interval so server won't get overloaded
                    Thread.Sleep(RandomGenerator.Next(3, 20) * 1000);
                    UpdateSignatures(true);
                });

                try
                {
                    _rxClient.ConnectAsync().Wait();
                }
                catch (Exception ex)
                {
                    _log.ErrorFormat("Couldn't connect to Rx Server: {0}", ex);
                }
            }
        }

        private void Addin_OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            _log.InfoFormat("Add-in unloading ({0})", RemoveMode);

            if (_rxClient != null)
                _rxClient.Dispose();
        }

        public void OnLoadRibonUI(Office.IRibbonUI ribbonUI)
        {
            RibbonUI = ribbonUI;
        }

        #endregion

        #region Static variables

        private static readonly string AppDataRoaming =
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

        private static IniData _cfg;
        private static ReactiveClient _rxClient;
        private static readonly Random RandomGenerator = new Random(DateTime.Now.Millisecond);

        private static ILog _log;
        private static Outlook.Application _myHost;
        private static Outlook.NameSpace _nsMapi;
        private static readonly string MyLocation = Assembly.GetExecutingAssembly().Location;
        private static IList<MailAddress> _mailAddresses;

        #endregion

        #region Ribbon button events

        public void AboutButton_Click(Office.IRibbonControl control)
        {
            var aboutForm = new AboutForm();
            aboutForm.ShowDialog();
            aboutForm.Dispose();
        }

        public void RefreshButton_Click(Office.IRibbonControl control)
        {
            if (IsNet45OrNewer())
                UpdateSignatures(true);
        }

        #endregion

        #region Addin error handler events

        protected override void OnError(ErrorMethodKind methodKind, Exception exception)
        {
            _log.ErrorFormat("An error occurend in {0}", methodKind);
            _log.ErrorFormat("Exception: {0}", exception);
        }

        [RegisterErrorHandler]
        public static void RegisterErrorHandler(RegisterErrorMethodKind methodKind, Exception exception)
        {
            _log.ErrorFormat("An error occurend in {0}", methodKind);
            _log.ErrorFormat("Exception: {0}", exception);
        }

        #endregion
    }
}