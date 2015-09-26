﻿using ExtensionMethods;
using PowerShellPowered.PowerShellConnect.Entities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Remoting;
using System.Management.Automation.Runspaces;
using System.Reflection;
using System.Security;
using System.Text;
using System.Timers;

namespace PowerShellPowered.PowerShellConnect
{
    public class PowerShellConnection : IDisposable
    {
        private Runspace _runspace;
        private Runspace _remoteSessionLocalRunspaceForDefultRunspace;
        public Runspace Runspace;
        private string ConnectionUri;
        private string SchemaUri;
        private string UserName;
        private SecureString Password;
        private int MaxRedirectionCount;
        //public PowerShellRunspaceMode RunspaceMode { get; private set; }
        public PowerShellRunspaceMode RunspaceMode { get; set; }
        public AuthenticationMechanism AauthenticationMechanism { get; private set; }
        internal ExecutionRunspace ExecutionRunspace;
        public bool RunspaceCreated { get; private set; }
        public string CerificateThumbPrint { get; private set; }
        public Guid CurrentPSInstanceID { get; set; }
        //public bool Ispersistant { get; set; }
        public List<string> ImportedModuleList { get; set; }
        public Action<RunspaceAvailability> RunspaceAvailabilityChanged { get; set; }
        public Action<RunspaceStateInfo> RunspaceStateChanged { get; set; }
        Timer CleanUpTimer;
        //public bool ImportPSSession { get; set; }

        public Dictionary<string, object> DefinedInitialVariables { get; set; }


        public PowerShellConnection(string connectionUri, string configSchemaUri, string userName, SecureString securePassword, PowerShellRunspaceMode runspaceMode = PowerShellRunspaceMode.RemoteRunspace, int maxRedirectionCount = 0, AuthenticationMechanism authenticationMechanism = AuthenticationMechanism.Basic)
            : this()
        {
            ConnectionUri = connectionUri;
            SchemaUri = configSchemaUri;
            UserName = userName;
            Password = securePassword;
            this.RunspaceMode = runspaceMode;
            this.AauthenticationMechanism = authenticationMechanism;
            MaxRedirectionCount = maxRedirectionCount;
        }
        public PowerShellConnection(string userName = "", SecureString securePassword = null, PowerShellRunspaceMode runspacemode = PowerShellRunspaceMode.LocalRunspace, AuthenticationMechanism authenticationMechanism = AuthenticationMechanism.NegotiateWithImplicitCredential)
        {
            this.RunspaceMode = runspacemode;
            this.AauthenticationMechanism = authenticationMechanism;
            UserName = userName;
            Password = securePassword;
            ImportedModuleList = new List<string>();
            DefinedInitialVariables = new Dictionary<string, object>();
        }
        public PowerShellConnection(string connectionUri, string configSchemaUri, string certificateThumbprint, PowerShellRunspaceMode runspaceMode = PowerShellRunspaceMode.RemoteRunspace, int maxRedirectionCount = 0, AuthenticationMechanism authenticationMechanism = AuthenticationMechanism.Credssp)
            : this()
        {
            ConnectionUri = connectionUri;
            SchemaUri = configSchemaUri;
            CerificateThumbPrint = certificateThumbprint;
            this.RunspaceMode = runspaceMode;
            this.AauthenticationMechanism = authenticationMechanism;
            MaxRedirectionCount = maxRedirectionCount;
        }
        ~PowerShellConnection()
        {
            RemoveRunspace();
        }
        public void Dispose()
        {
            RemoveRunspace();
        }

        public void Create(bool force = false, Action<PSDataStreams> psDataStreamAction = null)
        {
            CleanUpTimer = new Timer(300000);
            CleanUpTimer.AutoReset = false;
            //CleanUpTimer.SynchronizingObject = this;
            CleanUpTimer.Elapsed += (o, e) => { ClenupUnusedSecondaryConnection(); };


            if (RunspaceCreated && !force) throw new RemoteException("Runspace already created, please create new instance of PowerShellConnection");
            else if (RunspaceCreated)
                RemoveRunspace();

            //if (!Ispersistant)
            //    return;

            CreateRunspace(psDataStreamAction);
        }

        internal void ClenupUnusedSecondaryConnection()
        {
            if (RunspaceMode == PowerShellRunspaceMode.RemoteRunspace && ExecutionRunspace.PSSession != null)
                ExecutionRunspace.Dispose();

        }

        private void AttachRunspaceChanged(Runspace runspace)
        {
            if (runspace == null || (RunspaceAvailabilityChanged == null && RunspaceStateChanged == null))
                return;

            if (RunspaceStateChanged != null)
            {
                runspace.StateChanged += (o, e) =>
                {
                    RunspaceStateChanged(e.RunspaceStateInfo);
                };
            }

            if (RunspaceAvailabilityChanged != null)
            {
                runspace.AvailabilityChanged += (o, e) =>
                {
                    RunspaceAvailabilityChanged(e.RunspaceAvailability);
                };
            }

        }

        private void RemoveRunspace()
        {
            try
            {
                if (ExecutionRunspace != null)
                    ExecutionRunspace.Dispose();
            }
            catch (Exception)
            { }
            try
            {
                if (Runspace != null)
                {
                    if (this.Runspace.RunspaceAvailability != RunspaceAvailability.None)
                        this.Runspace.Close();
                    this.Runspace.Dispose();
                }
            }
            catch (Exception)
            { }
            RunspaceCreated = false;
        }

        public void ResetRunspace()
        {
            RemoveRunspace();
            Create(true);

        }


        private WSManConnectionInfo GetWsManConnection()
        {
            PSCredential cred = new PSCredential(UserName, Password);
            WSManConnectionInfo connectionInfo = null;
            switch (AauthenticationMechanism)
            {
                case AuthenticationMechanism.Basic:
                    connectionInfo = new WSManConnectionInfo(new Uri(ConnectionUri), SchemaUri, cred);
                    //connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;
                    break;
                case AuthenticationMechanism.Credssp:
                    connectionInfo = new WSManConnectionInfo(new Uri(ConnectionUri), SchemaUri, CerificateThumbPrint);
                    //connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Credssp;
                    break;
                case AuthenticationMechanism.Kerberos:
                    connectionInfo = new WSManConnectionInfo(new Uri(ConnectionUri), SchemaUri, cred);
                    //connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Kerberos;
                    break;
                case AuthenticationMechanism.Negotiate:
                    connectionInfo = new WSManConnectionInfo(new Uri(ConnectionUri), SchemaUri, cred);
                    //connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Negotiate;
                    break;
                case AuthenticationMechanism.NegotiateWithImplicitCredential:
                    connectionInfo = new WSManConnectionInfo();
                    break;
                default:
                    throw new NotImplementedException("Unknown or not implemented Authentication Mechanism");
            }
            connectionInfo.AuthenticationMechanism = AauthenticationMechanism;
            return connectionInfo;
        }


        #region Reference_from_msdn
        private Collection<PSObject> GetUsersUsingBasicAuth(
            string liveIDConnectionUri, string schemaUri, PSCredential credentials, int count, string cmd)
        {
            WSManConnectionInfo connectionInfo = new WSManConnectionInfo(
                new Uri(liveIDConnectionUri),
                schemaUri, credentials);
            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;

            using (Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo))
            {
                //return GetUserInformation(count, runspace);
                //return GetCommandOutput(count, runspace, cmd);
                return null;
            }
        }

        private Collection<PSObject> GetUsersUsingCertificate(
            string thumbprint, string certConnectionUri, string schemaUri, int count)
        {
            WSManConnectionInfo connectionInfo = new WSManConnectionInfo(
                new Uri(certConnectionUri),
                schemaUri,
                thumbprint);

            using (Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo))
            {
                return GetUserInformation(count, runspace);
            }
        }

        private Collection<PSObject> GetUsersUsingKerberos(
            string kerberosUri, string schemaUri, PSCredential credentials, int count)
        {
            WSManConnectionInfo connectionInfo = new WSManConnectionInfo(
                new Uri(kerberosUri),
                schemaUri, credentials);
            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Kerberos;

            using (Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo))
            {
                return GetUserInformation(count, runspace);
            }
        }


        private Collection<PSObject> GetUsersUsingNegotiatedAuth(string schemaUri, string targetUri, int count, PSCredential credential)
        {
            WSManConnectionInfo connectionInfo = new WSManConnectionInfo(
              new Uri(targetUri),
              schemaUri, credential);

            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Negotiate;

            using (Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo))
            {
                return GetUserInformation(count, runspace);
            }
        }

        private Collection<PSObject> GetUserInformation(int count, Runspace runspace)
        {
            using (PowerShell powershell = PowerShell.Create())
            {

                //powershell.AddCommand("Get-Recipient");
                //if(count==0)
                //    powershell.AddParameter("ResultSize", "unlimited");
                //else
                //    powershell.AddParameter("ResultSize", count);

                powershell.AddCommand("Get-command");



                runspace.Open();


                powershell.Runspace = runspace;
                Collection<PSObject> _result = powershell.Invoke();

                return _result;
            }
        }

        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="runspace"></param>
        /// <param name="command"></param>
        /// <param name="parameters"> Dictionary of string parametername, KeyValue -Parametertype, param value </param>
        /// <returns></returns>
        public Collection<PSObject> ExecuteCommand(string command, Dictionary<string, object> parameters, bool useSession = false, Action<PSDataStreams> psDataStreamAction = null)
        {
            return ExecuteCommand<PSObject>(command, parameters, false, psDataStreamAction);

        }

        public Collection<T> ExecuteCommand<T>(string command, Dictionary<string, object> parameters, bool useSession = false, Action<PSDataStreams> psDataStreamAction = null)
        {
            Command cmd = new Command(command);

            if (parameters != null)
                foreach (KeyValuePair<string, object> item in parameters)
                {
                    if (item.Value == null) cmd.Parameters.Add(item.Key);
                    else cmd.Parameters.Add(item.Key, item.Value);
                }
            PSCommand psc = new PSCommand();
            psc.AddCommand(cmd);

            _runspace = GetExecutionRunspace();

            var result = _runspace.ExecuteCommand<T>(psc, psDataStreamAction);

            //if (!Ispersistant) ResetRunspace();

            return result;
        }

        public Collection<PSObject> ExecuteCommand(PSCommand command, Action<PSDataStreams> psDataStreamAction = null)
        {
            //_runspace = GetExecutionRunspace();
            //return _runspace.ExecuteCommand<PSObject>(command, psDataStreamAction);
            return this.ExecuteCommand<PSObject>(command, psDataStreamAction);
        }

        public Collection<T> ExecuteCommand<T>(PSCommand command, Action<PSDataStreams> psDataStreamAction = null)
        {
            _runspace = GetExecutionRunspace();

            var result = _runspace.ExecuteCommand<T>(command, psDataStreamAction);

            //if (!Ispersistant) ResetRunspace();

            return result;
        }

        public Collection<PSObject> ExecuteScript(string scriptText, Dictionary<string, object> scriptParameters = null, bool useSession = true, bool out_String = false, bool stream = false, int width = 10000, Action<PSDataStreams> psDataStreamAction = null)
        {
            return this.ExecuteScript<PSObject>(scriptText, scriptParameters, useSession, out_String, stream, width, psDataStreamAction);

            //return this.ExecutePipeLine<PSObject>(scriptText, out_String, stream, width, psDataStreamAction);
        }

        public Collection<T> ExecuteScript<T>(string scriptText, Dictionary<string, object> scriptParameters = null, bool useSession = true, bool out_String = false, bool stream = false, int width = 10000, Action<PSDataStreams> psDataStreamAction = null)
        {
            _runspace = GetExecutionRunspace();

            if (useSession)
            {
                if (scriptParameters == null)
                    scriptParameters = new Dictionary<string, object>();
                if (ExecutionRunspace != null && ExecutionRunspace.PSSession != null)
                    if (scriptParameters.ContainsKey("PsSession") && scriptParameters["PsSession"] != null)
                        scriptParameters["PsSession"] = ExecutionRunspace.PSSession;
                //else
                //    scriptParameters.Add("Session", ExecutionRunspace.PSSession);
            }

            //? using runspace extension method to simplify management of code.
            var result = _runspace.ExecuteScript<T>(scriptText, scriptParameters, out_String, stream, width, psDataStreamAction);

            //if (!Ispersistant) ResetRunspace();

            return result;
        }

        public void SetDefaultRunspace()
        {
            //System.Management.Automation.Runspaces.Runspace.DefaultRunspace = GetExecutionRunspace();
            // above line useless as getexecutionrunspace set default runspace everytime.
            GetExecutionRunspace();
        }

        public void SetVariable(string variableName, object variableValue)
        {
            _runspace = GetExecutionRunspace();
            _runspace.SetRunspaceVariable(variableName, variableValue);

            if (DefinedInitialVariables.ContainsKey(variableName)) DefinedInitialVariables[variableName] = variableValue;
            else DefinedInitialVariables.Add(variableName, variableValue);
        }

        void SetupdefinedVariables(Runspace runspace)
        {
            if (runspace == null || runspace.RunspaceStateInfo.State != RunspaceState.Opened) return;
            foreach (var item in DefinedInitialVariables)
            {
                runspace.SetRunspaceVariable(item.Key, item.Value);
            }
        }

        public void ImportModules(string[] modules, Action<PSDataStreams> psDataStreamAction = null)
        {
            if (modules == null || modules.Length == 0) return;

            Dictionary<string, object> cmdparams = new Dictionary<string, object>();
            cmdparams.Add("Name", modules);
            var runspace = GetExecutionRunspace();
            ExecuteCommand("Import-Module", cmdparams, false, psDataStreamAction);
            modules.ToList().ForEach((x) => { if (!ImportedModuleList.Contains(x)) { ImportedModuleList.Add(x); } });
        }

        public void RemoveModules(string[] modules, Action<PSDataStreams> psDataStreamAction = null)
        {
            if (modules == null || modules.Length == 0) return;

            Dictionary<string, object> cmdparams = new Dictionary<string, object>();
            cmdparams.Add("Name", modules);
            var runspace = GetExecutionRunspace();
            ExecuteCommand("Remove-Module", cmdparams, false, psDataStreamAction);
            modules.ToList().ForEach((x) => { if (ImportedModuleList.Contains(x)) { ImportedModuleList.Remove(x); } });
        }

        public Collection<PSModuleInfoProxy> GetModuleInformation(string[] modules = null, bool listavailable = false, Action<PSDataStreams> psDataStreamAction = null)
        {
            Dictionary<string, object> cmdparams = new Dictionary<string, object>();
            if (modules != null) cmdparams.Add("Name", modules);
            if (listavailable) cmdparams.Add("ListAvailable", null);

            Collection<PSModuleInfoProxy> result = new Collection<PSModuleInfoProxy>();

            result = GetModuleInfoProxyCollection(ExecuteCommand<PSModuleInfo>("Get-Module", cmdparams, false, psDataStreamAction));

            return result;
        }

        private Collection<PSModuleInfoProxy> GetModuleInfoProxyCollection(Collection<PSModuleInfo> moduleinfoInfoCollection)
        {
            Collection<PSModuleInfoProxy> _result = new Collection<PSModuleInfoProxy>();
            foreach (PSModuleInfo moduleinfo in moduleinfoInfoCollection)
            {
                _result.Add(moduleinfo.ToPSModuleInfoProxy());

                //try
                //{
                //}
                //catch (Exception ex)
                //{

                //}
            }
            return _result;
        }

        public Collection<PSCommandInfoProxy> GetCommandCollection(string[] cmdletArray = null, int totalCount = -1, string[] module = null, bool skipmodule = false, PowerShellCommandTypes commandType = PowerShellCommandTypes.AllPowerShellNative, string nameFilter = null, Action<PSDataStreams> psDataStreamAction = null, Action<PSDataAlternateStream> alternateDataStreamAction = null)
        {
            Dictionary<string, object> cmdparams = new Dictionary<string, object>();
            //if (module == null && RunspaceMode == PowerShellRunspaceMode.LocalRunspace)
            //    module = new string[] { "Microsoft.PowerShell.Core" };

            if (RunspaceMode == PowerShellRunspaceMode.LocalRunspace && !skipmodule)
            {
                if (module != null) cmdparams.Add(Constants.PSParameterNameStrings.Module, module);
                else if (ImportedModuleList != null && ImportedModuleList.Count > 0) cmdparams.Add(Constants.PSParameterNameStrings.Module, ImportedModuleList.ToArray());
            }

            if (!nameFilter.IsNullOrEmpty()) cmdparams.Add(Constants.PSParameterNameStrings.Name, nameFilter);
            else if (cmdletArray != null) cmdparams.Add(Constants.PSParameterNameStrings.Name, cmdletArray);
            cmdparams.Add(Constants.PSParameterNameStrings.CommandType, (int)commandType);
            cmdparams.Add("TotalCount", totalCount);

            cmdparams.Add("listimported", null);

            Collection<PSCommandInfoProxy> result = new Collection<PSCommandInfoProxy>();
            if (RunspaceMode == PowerShellRunspaceMode.LocalRunspace)
            {
                Command cmd = new Command(Constants.PSCmdlets.GetCommand);
                if (cmdparams != null)
                    foreach (KeyValuePair<string, object> item in cmdparams)
                    {
                        if (item.Value == null) cmd.Parameters.Add(item.Key);
                        else cmd.Parameters.Add(item.Key, item.Value);
                    }
                PSCommand psc = new PSCommand();
                psc.AddCommand(cmd);
                psc.AddCommand("where-object").AddParameter("property", "modulename").AddParameter("ne").AddParameter("value", "");
                //result = GetCommandInfoProxyCollection(ExecuteCommand<CommandInfo>(Constants.PSCmdlets.GetCommand, cmdparams, false, psDataStreamAction), alternateDataStreamAction);
                result = GetCommandInfoProxyCollection(ExecuteCommand<CommandInfo>(psc, psDataStreamAction), alternateDataStreamAction);
            }

            else
            {
                result = GetCommandCollectionRemote(cmdletArray, totalCount, module, commandType, nameFilter, psDataStreamAction);
            }

            //if (!Ispersistant) ResetRunspace();

            return result;
        }

        public Collection<PSCommandInfoProxy> GetScriptInfo(string scriptFileName, PowerShellCommandTypes commandType = PowerShellCommandTypes.ExternalScript, Action<PSDataStreams> psDataStreamAction = null, Action<PSDataAlternateStream> alternateDataStreamAction = null)
        {
            Dictionary<string, object> cmdparams = new Dictionary<string, object>();


            cmdparams.Add(Constants.PSParameterNameStrings.Name, scriptFileName);
            //cmdparams.Add(Constants.PSParameterNameStrings.CommandType, (int)commandType);

            Collection<PSCommandInfoProxy> result = new Collection<PSCommandInfoProxy>();
            result = GetCommandInfoProxyCollection(ExecuteCommand<CommandInfo>(Constants.PSCmdlets.GetCommand, cmdparams, false, psDataStreamAction), alternateDataStreamAction);

            //if (!Ispersistant) ResetRunspace();

            return result;
        }

        public Collection<PSObject> GetCommandCollectionRaw(string[] cmdletArray = null, int totalCount = -1, string[] module = null, PowerShellCommandTypes commandType = PowerShellCommandTypes.Cmdlet, string nameFilter = null)
        {
            Dictionary<string, object> cmdparams = new Dictionary<string, object>();
            if (module != null) cmdparams.Add(Constants.PSParameterNameStrings.Module, module);
            else if (ImportedModuleList != null && ImportedModuleList.Count > 0) cmdparams.Add(Constants.PSParameterNameStrings.Module, ImportedModuleList.ToArray());

            if (!nameFilter.IsNullOrEmpty()) cmdparams.Add(Constants.PSParameterNameStrings.Name, nameFilter);
            else if (cmdletArray != null) cmdparams.Add(Constants.PSParameterNameStrings.Name, cmdletArray);
            cmdparams.Add(Constants.PSParameterNameStrings.CommandType, (int)commandType);
            cmdparams.Add("TotalCount", totalCount);

            var result = ExecuteCommand<PSObject>(Constants.PSCmdlets.GetCommand, cmdparams);
            return result;

        }

        public Collection<PSObject> GetCommandCollectionScriptedRaw(string[] cmdletArray = null, int totalCount = -1, string[] module = null, PowerShellCommandTypes commandType = PowerShellCommandTypes.Cmdlet, string nameFilter = null)
        {
            Dictionary<string, object> cmdparams = new Dictionary<string, object>();
            if (module != null) cmdparams.Add(Constants.PSParameterNameStrings.Module, module);
            else if (ImportedModuleList != null && ImportedModuleList.Count > 0) cmdparams.Add(Constants.PSParameterNameStrings.Module, ImportedModuleList.ToArray());

            if (!nameFilter.IsNullOrEmpty()) cmdparams.Add(Constants.PSParameterNameStrings.Name, nameFilter);
            else if (cmdletArray != null) cmdparams.Add(Constants.PSParameterNameStrings.Name, cmdletArray);
            cmdparams.Add(Constants.PSParameterNameStrings.CommandType, (int)commandType);
            cmdparams.Add("TotalCount", totalCount);

            var result = ExecuteCommand<PSObject>(Constants.PSCmdlets.GetCommand, cmdparams);
            return result;

        }

        //CommandTypes
        //Alias
        //1
        //Function
        //2
        //Filter
        //4
        //Cmdlet
        //8
        //ExternalScript
        //16
        //Application
        //32
        //Script
        //64
        //Workflow
        //128
        //All
        //255

        private Collection<PSCommandInfoProxy> GetCommandCollectionRemote(string[] cmdletArray = null, int totalCount = -1, string[] module = null, PowerShellCommandTypes commandType = PowerShellCommandTypes.Cmdlet, string nameFilter = null, Action<PSDataStreams> psDataStreamAction = null)
        {

            if (RunspaceMode == PowerShellRunspaceMode.RemoteRunspace)
            {
                CleanUpTimer.Start();
                //GC.KeepAlive(CleanUpTimer);
                if (ExecutionRunspace == null || ExecutionRunspace.Runspace.RunspaceStateInfo.State != RunspaceState.Opened || ExecutionRunspace.PSSession.Runspace.RunspaceStateInfo.State != RunspaceState.Opened)
                {
                    this.ExecutionRunspace = new ExecutionRunspace(UserName, Password, ConnectionUri, SchemaUri, true, false, ImportedModuleList, psDataStreamAction);
                }
                _runspace = ExecutionRunspace.Runspace;
            }
            else
            {
                _runspace = GetExecutionRunspace();
            }


            Command cmd = new Command(Constants.PSHelpParsingScripts.GetCommandScript, true);
            if (RunspaceMode != PowerShellRunspaceMode.LocalRunspace)
                cmd.Parameters.Add("Session", ExecutionRunspace.PSSession);

            string assemblypath = Assembly.GetAssembly(typeof(PowerShellPowered.PowerShellConnect.PowerShellConnection)).Location;
            cmd.Parameters.Add("EntitiesFile", assemblypath);

            if (!nameFilter.IsNullOrEmpty()) cmd.Parameters.Add("Name", nameFilter);
            else if (cmdletArray != null) cmd.Parameters.Add("Name", cmdletArray);

            cmd.Parameters.Add("CommandType", (int)commandType);
            cmd.Parameters.Add("TotalCount", totalCount);
            if (RunspaceMode == PowerShellRunspaceMode.LocalRunspace)
                if (module != null) cmd.Parameters.Add("Module", module);
            PSCommand pscc = new PSCommand();
            pscc.Commands.Add(cmd);

            var _result = _runspace.ExecuteCommand<PSCommandInfoProxy>(pscc, psDataStreamAction);

            return _result;
        }

        private Collection<PSCommandInfoProxy> GetCommandInfoProxyCollection(Collection<CommandInfo> commandInfoCollection, Action<PSDataAlternateStream> alternateDataStream = null)
        {
            Collection<PSCommandInfoProxy> _result = new Collection<PSCommandInfoProxy>();
            int count = 0;
            int total = commandInfoCollection.Count;
            foreach (CommandInfo commandinfo in commandInfoCollection)
            {
                count++;
                int percent = (100 * count / total);
                if (alternateDataStream != null)
                    alternateDataStream(new PSDataAlternateStream() { HasProgress = true, SimpleProgress = new ProgressRecord(100, "adding data", "Adding Data Description") { PercentComplete = percent } });
                try
                {
                    _result.Add(commandinfo.ToPSCommandInfoProxy(true));
                }
                catch (Exception ex)
                {
                    if (alternateDataStream != null)
                    {
                        string errline = string.Format("Error parsing '{0}' in module '{1}'. Message: {2}\r\n{3}", commandinfo.Name, commandinfo.ModuleName, ex.Message, (ex.InnerException != null ? ex.InnerException.Message : string.Empty));
                        alternateDataStream(new PSDataAlternateStream() { HasSimpleError = true, SimpleError = errline });
                    }
                }
            }
            if (alternateDataStream != null)
                alternateDataStream(new PSDataAlternateStream() { HasProgress = true, SimpleProgress = new ProgressRecord(100, "adding data", "Adding Data Description") { PercentComplete = 100, RecordType = ProgressRecordType.Completed } });
            return _result;
        }

        public static string GetRedirectedURL(string connectionUri, string configSchemaUri, string userName, SecureString securePassword, PowerShellRunspaceMode runspaceMode = PowerShellRunspaceMode.RemoteRunspace, int maxRedirectionCount = 0)
        {

            string _r = connectionUri;
            Dictionary<string, object> _params = new Dictionary<string, object>();
            _params.Add(Constants.PSParameterNameStrings.Name, Constants.PSCmdlets.GetCommand);
            try
            {
                using (PowerShellConnection ps = new PowerShellConnection(_r, configSchemaUri, userName, securePassword, maxRedirectionCount: maxRedirectionCount))
                {
                    ps.Create();
                }
            }
            catch (PSRemotingTransportRedirectException ex)
            {
                return ex.RedirectLocation;
                //_result = powershell.Invoke();
            }
            catch (Exception)
            {

                throw;
            }

            return _r;

        }

        private string GetCommandSynopsis(string command, string psModule = null)
        {
            PSCommand psc = new PSCommand();
            psc.AddCommand(Constants.PSCmdlets.GetHelp).AddParameter(Constants.PSParameterNameStrings.Name, command);
            try
            {
                PSObject pso = GetExecutionRunspace().ExecuteCommand(psc)[0];
                return pso.GetPropertyValue("Synopsis");
            }
            catch (Exception)
            { }

            return string.Empty;
        }

        public PSCommandHelp GetCommandHelp(string command, Action<PSDataStreams> psDataStreamAction)
        {
            if (RunspaceMode == PowerShellRunspaceMode.RemoteRunspace)
            {
                CleanUpTimer.Start();
                //GC.KeepAlive(CleanUpTimer);
                if (ExecutionRunspace == null || ExecutionRunspace.Runspace.RunspaceStateInfo.State != RunspaceState.Opened || ExecutionRunspace.PSSession.Runspace.RunspaceStateInfo.State != RunspaceState.Opened)
                {
                    this.ExecutionRunspace = new ExecutionRunspace(UserName, Password, ConnectionUri, SchemaUri, true, false, ImportedModuleList, psDataStreamAction);
                }
                else
                {
                    _runspace = ExecutionRunspace.Runspace;
                }
            }
            else
            {
                _runspace = GetExecutionRunspace();
            }

            PSCommand pscmd = new PSCommand();
            //string script = string.Format(Constants.PSHelpParsingScripts.GetHelpScript, command);
            //Command cmd = new Command(script, true);
            Command cmd = new Command(Constants.PSHelpParsingScripts.GetHelpScript, true);
            cmd.Parameters.Add("Name", command);
            string assemblypath = Assembly.GetAssembly(typeof(PowerShellConnection)).Location;
            cmd.Parameters.Add("EntitiesFile", assemblypath);

            if (RunspaceMode != PowerShellRunspaceMode.LocalRunspace)
                cmd.Parameters.Add("Session", ExecutionRunspace.PSSession);

            pscmd.AddCommand(cmd);
            PSCommandHelp hcmd = null;
            try
            {
                hcmd = new PSCommandHelp();

                Collection<PSCommandHelp> res = _runspace.ExecuteCommand<PSCommandHelp>(pscmd, psDataStreamAction);
                hcmd = res[0];
                //! using simplified methods with script text
                //++ win8/psv3 has bug for get-help -full, run update-help to stage local help                
                //cmd.AddCommand(Constants.PSCmdlets.GetHelp).AddParameter(Constants.PSParameterNameStrings.Name, command);//.AddParameter("Full");
                //Collection<PSObject> res = ExecutionRunspaceLocal.ExecuteCommand(cmd);

                //++ skipped for test
                //////hcmd = GetCommandHelp(res[0]);

                //////string script = "get-help get-mailbox";
                //////var scrc = new Command(script, true);

                //////script = "(get-help get-mailbox).parameters[0].parameter";
                //////scrc = new Command(script, true);
                //////cmd = new PSCommand();
                //////cmd.AddCommand(script);
                //////var xc = ExecutionRunspaceLocal.ExecutePipeLine(script, true);
                //////var cv = ExecutionRunspaceLocal.ExecuteCommand(cmd);

                //////var reader = new System.IO.StreamReader(@"D:\Temp DLs\Get_parameterInfo.ps1");
                //////script = reader.ReadToEnd();
                ////////script = "$g = Get-Help Get-Mailbox; $g.parameters;";
                //////scrc = new Command(script, true);
                //////cmd = new PSCommand();
                //////cmd.AddCommand(scrc);

                //////cv = ExecutionRunspaceLocal.ExecuteCommand(cmd);
                //////var y = cv;

                //////try
                //////{
                //////    var bcv = cv[8].Properties;
                //////    foreach (var item in bcv)
                //////    {
                //////        var bvc = item.Value;
                //////    }
                //////}
                //////catch (Exception exx)
                //////{
                //////    var exxx = exx;

                //////}


                //hcmd.Name = ExecutionMethods.FormatPowerShellCommandName(command);
                //////hcmd.Synopsis = ExecutionRunspaceLocal.ExecutePipeLine(string.Format(Constants.PSHelpParsingScripts.SynopsisScript, command), true).RemoveLeadingNewLine();
                //hcmd.Details = ExecutionRunspaceLocal.ExecutePipeLine(string.Format(Constants.PSHelpParsingScripts.Detailscript, command), true).Trim();
                //hcmd.Syntax = ExecutionRunspaceLocal.ExecutePipeLine(string.Format(Constants.PSHelpParsingScripts.SyntaxScript, command), true).Trim();
                //hcmd.Description = ExecutionRunspaceLocal.ExecutePipeLine(string.Format(Constants.PSHelpParsingScripts.DescriptionScript, command), true).Trim();

                ////!++ win8/psv3 has bug for get-help -full fix constants.cs after bug is fixed
                //hcmd.Parameters = ExecutionRunspaceLocal.ExecutePipeLine(string.Format(Constants.PSHelpParsingScripts.ParametersScript, command), true).Trim();
                //hcmd.InputTypes = ExecutionRunspaceLocal.ExecutePipeLine(string.Format(Constants.PSHelpParsingScripts.InputtypesScript, command), true).Trim();
                //hcmd.ReturnValues = ExecutionRunspaceLocal.ExecutePipeLine(string.Format(Constants.PSHelpParsingScripts.ReturnValuesScript, command), true).Trim();
                //hcmd.TerminatingErrors = ExecutionRunspaceLocal.ExecutePipeLine(string.Format(Constants.PSHelpParsingScripts.TerminatingErrorsScript, command), true).Trim();
                //hcmd.NonTerminatingErrors = ExecutionRunspaceLocal.ExecutePipeLine(string.Format(Constants.PSHelpParsingScripts.NonTerminatingErrorsScript, command), true).Trim();
                //hcmd.Examples = ExecutionRunspaceLocal.ExecutePipeLine(string.Format(Constants.PSHelpParsingScripts.ExamplesScript, command), true).Trim();
                //hcmd.RelatedLinks = ExecutionRunspaceLocal.ExecutePipeLine(string.Format(Constants.PSHelpParsingScripts.RelatedLinksScript, command), true).Trim();
                //////hcmd.Category = ExecutionRunspaceLocal.ExecutePipeLine(string.Format(Constants.PSHelpParsingScripts.RelatedLinksScript, command), true).RemoveLeadingNewLine();
                //////hcmd.Component = ExecutionRunspaceLocal.ExecutePipeLine(string.Format(Constants.PSHelpParsingScripts.ComponentScript, command), true).RemoveLeadingNewLine();
                //////hcmd.Role = ExecutionRunspaceLocal.ExecutePipeLine(string.Format(Constants.PSHelpParsingScripts.RoleScript, command), true).RemoveLeadingNewLine();
                //////hcmd.Functionality = ExecutionRunspaceLocal.ExecutePipeLine(string.Format(Constants.PSHelpParsingScripts.FunctionalityScript, command), true).RemoveLeadingNewLine();
                //hcmd.FullHelpText = ExecutionRunspaceLocal.ExecutePipeLine(string.Format(Constants.PSHelpParsingScripts.FullHelpScript, command), true).Trim();
            }
            catch (Exception ex)
            {

                throw new Exception("thrown in gethelp execute local" + ex.Message);
            }

            //if (RunspaceMode != originalRunspaceMode)
            //{
            //    RunspaceMode = originalRunspaceMode;
            //    if (Ispersistant) Create();
            //}

            return hcmd;
        }


        #region Get Command Help Private Methods

        private static PSCommandHelp GetCommandHelp(PSObject pso)
        {
            PSCommandHelp _result = new PSCommandHelp();
            //--------alternate test

            _result.Category = pso.GetPropertyValue("category");
            _result.Component = pso.GetPropertyValue("component");
            _result.Description = pso.GetPropertyValue("description");
            _result.Details = pso.GetPropertyValue("details");
            _result.Examples = pso.GetPropertyValue("examples");
            _result.FullHelpText = pso.GetPropertyValue("fullhelptext");
            _result.Functionality = pso.GetPropertyValue("functionality");
            _result.InputTypes = pso.GetPropertyValue("inputtypes");
            _result.Name = pso.GetPropertyValue("name");
            _result.NonTerminatingErrors = pso.GetPropertyValue("nonterminatingerrors");
            _result.Parameters = pso.GetPropertyValue("parameters");
            _result.RelatedLinks = pso.GetPropertyValue("relatedlinks");
            _result.ReturnValues = pso.GetPropertyValue("returnvalues");
            _result.Role = pso.GetPropertyValue("role");
            _result.Synopsis = pso.GetPropertyValue("synopsis");
            _result.Syntax = pso.GetPropertyValue("syntax");
            _result.TerminatingErrors = pso.GetPropertyValue("terminatingerrors");
            PSObject[] parameters;
            pso.TryGetPropertyAsTObject<PSObject[]>("ParametersEx", out parameters);
            //PSObject p = new PSObject(px);
            //pso.TryGetPropertyAsPSObject("ParametersEx", out parameters);
            _result.PSCommandHelpParameters = new List<PSCommandHelpParameter>();
            foreach (var p in parameters)
            {
                _result.PSCommandHelpParameters.AddRange(GetCommandParameters(p));
            }

            PSObject[] examples;
            pso.TryGetPropertyAsTObject<PSObject[]>("ExamplesEx", out examples);
            //PSObject p = new PSObject(px);
            //pso.TryGetPropertyAsPSObject("ParametersEx", out parameters);
            _result.PSCommandHelpExamples = new List<PSCommandHelpExample>();
            foreach (var e in examples)
            {
                _result.PSCommandHelpExamples.AddRange(GetCommandExample(e));
            }

            PSObject[] syntaxes;
            pso.TryGetPropertyAsTObject<PSObject[]>("SyntaxEx", out syntaxes);
            //PSObject p = new PSObject(px);
            //pso.TryGetPropertyAsPSObject("ParametersEx", out parameters);
            _result.PSCommandHelpSyntaxes = new List<PSCommandHelpSyntax>();
            foreach (var s in syntaxes)
            {
                _result.PSCommandHelpSyntaxes.AddRange(GetCommandSyntax(s));
            }



            // -------/alternate test
            ////_result.Id = Guid.NewGuid();

            //StringBuilder sb = new StringBuilder();
            //_result.Name = pso.GetPropertyValue("Name");
            //if (string.IsNullOrEmpty(_result.Name)) throw new Exception("Name can not be blank");

            //_result.Category = pso.GetPropertyValue("Category");
            //_result.Synopsis = pso.GetPropertyValue("Synopsis");
            //_result.Component = pso.GetPropertyValue("Component");
            //_result.Role = pso.GetPropertyValue("Role");
            //_result.Functionality = pso.GetPropertyValue("Functionality");

            //PSObject psParameters;
            //if (pso.TryGetPropertyAsPSObject("parameters", out psParameters))
            //{
            //    PSObject psParameter;
            //    if (psParameters.Properties.Match("parameter").Count > 0)
            //    {
            //        var psParameterList = psParameters.Properties["parameter"].Value as IList;
            //        if (psParameterList != null)
            //        {
            //            foreach (var item in psParameterList)
            //            {
            //                _result.PSCmdletHelpParameters.AddRange(GetCommandParameters(item as PSObject));
            //            }
            //        }
            //        else throw new ItemNotFoundException("unable to get parameter list");
            //    }
            //    else if (psParameters.ImmediateBaseObject is PSObject && ((PSObject)psParameters.ImmediateBaseObject).TryGetPropertyAsPSObject("parameter", out psParameter))
            //    {
            //        _result.PSCmdletHelpParameters = GetCommandParameters(psParameter);
            //    }
            //}

            //PSObject psExamples;
            //if (pso.TryGetPropertyAsPSObject("Examples", out psExamples))
            //{
            //    PSObject psExample;
            //    if (psExamples.ImmediateBaseObject is PSObject && ((PSObject)psExamples.ImmediateBaseObject).TryGetPropertyAsPSObject("Example", out psExample))
            //    {
            //        _result.PSCmdletHelpExamples = GetCommandExample(psExample);
            //    }
            //}

            //PSObject psSyntax;
            //if (pso.TryGetPropertyAsPSObject("Syntax", out psSyntax))
            //{
            //    PSObject psSyntaxItem;
            //    if (psSyntax.ImmediateBaseObject is PSObject && ((PSObject)psSyntax.ImmediateBaseObject).TryGetPropertyAsPSObject("SyntaxItem", out psSyntaxItem))
            //    {
            //        _result.PSCmdletHelpSyntaxes = GetCommandSyntax(psSyntaxItem);
            //    }
            //}

            return _result;

        }
        private static List<PSCommandHelpParameter> GetCommandParameters(PSObject pso)
        {
            List<PSCommandHelpParameter> _result = new List<PSCommandHelpParameter>();

            if (pso.IsArrayList())
            {
                foreach (PSObject parameter in pso.ToArrayList())
                {
                    List<PSCommandHelpParameter> _r_nested = GetCommandParameters(parameter);
                    if (_r_nested != null) _result.AddRange(_r_nested);
                }
            }
            else
            {
                PSCommandHelpParameter p = new PSCommandHelpParameter();
                StringBuilder sb = new StringBuilder();
                //p.Id = Guid.NewGuid();
                //p.CommandHelpId = commandId;
                p.Name = pso.GetPropertyValue("Name");
                if (string.IsNullOrEmpty(p.Name)) return null;

                p.Description = pso.GetPropertyValue("description");
                //p.Description = GetCommandDescription(pso);
                if (string.IsNullOrEmpty(p.Description))
                    p.Description = "No Description or error parsing Description";
                p.ParameterValue = pso.GetPropertyValue("ParameterValue");
                p.Position = pso.GetPropertyValue("Position");
                p.Required = pso.GetPropertyValue("Required").ParseBool();
                p.PipelineInput = pso.GetPropertyValue("PipelineInput").ParseBool();
                p.VariableLength = pso.GetPropertyValue("VariableLength").ParseBool();
                p.Globbing = pso.GetPropertyValue("Globbing").ParseBool();
                p.DefaultValue = pso.GetPropertyValue("DefaultValue");

                p.Type = pso.GetPropertyValue("type");
                //? not needed with script based extraction
                //PSObject psType;
                //if (pso.TryGetPropertyAsPSObject("Type", out psType))
                //{
                //    if (psType.IsArrayList())
                //    {
                //        p.Type = "Paramtype is ArrayList. check and fix this line of code in getcommandparameter";
                //    }
                //    else
                //    {
                //        p.Type = psType.GetPropertyValue("Name");
                //    }
                //}
                //else
                //{
                //    p.Type = pso.GetPropertyValue("Type");
                //}
                _result.Add(p);
            }

            return _result;
        }

        private static string GetCommandDescription(PSObject pso)
        {
            StringBuilder sb = new StringBuilder();
            IList psDescriptionIlist;

            if (pso.TryGetPropertyAsTObject<IList>("description", out psDescriptionIlist))
            {
                foreach (var item in psDescriptionIlist)
                {
                    if (item is PSObject) sb.AppendLine(GetCommandDescriptionTex(item as PSObject));
                }
                return sb.ToString();
            }

            PSObject psDescription;
            if (pso.TryGetPropertyAsPSObject("Description", out psDescription))
            {
                if (psDescription.IsArrayList())
                    foreach (PSObject desc in psDescription.ToArrayList())
                    {
                        sb.AppendLine(GetCommandDescriptionTex(desc));
                    }
                else
                {
                    sb.AppendLine(GetCommandDescriptionTex(psDescription));
                }
            }

            return sb.ToString();
        }
        private static string GetCommandDescriptionTex(PSObject pso)
        {
            string descriptionLine = string.Empty;
            string Tag = pso.GetPropertyValue("Tag");
            string Text = pso.GetPropertyValue("Text");

            if (!string.IsNullOrEmpty(Tag)) descriptionLine += Tag;
            if (!string.IsNullOrEmpty(Text)) descriptionLine += Text;
            return descriptionLine;
        }

        private static List<PSCommandHelpExample> GetCommandExample(PSObject pso)
        {
            List<PSCommandHelpExample> _result = new List<PSCommandHelpExample>();

            if (pso.IsArrayList())
            {
                foreach (PSObject example in pso.ToArrayList())
                {
                    List<PSCommandHelpExample> _r_nested = GetCommandExample(example);
                    if (_r_nested != null) _result.AddRange(_r_nested);
                }
            }
            else
            {


                PSCommandHelpExample example = new PSCommandHelpExample();
                StringBuilder sb = new StringBuilder();
                //example.Id = Guid.NewGuid();
                //example.CommandHelpId = commandId;

                example.Title = pso.GetPropertyValue("Title");
                if (example.Title.IsNullOrEmpty()) new Exception("Example Title can not be null");
                example.Title = example.Title.ToString().Replace("-", "").Trim();

                example.Code = pso.GetPropertyValue("Code");
                example.Introduction = pso.GetPropertyValue("Introduction");

                //? using new scripted methods to extract all text
                example.CommandLines = pso.GetPropertyValue("CommandLines");
                example.Remarks = pso.GetPropertyValue("Remarks");

                //PSObject psCmdLines;
                //if (pso.TryGetPropertyAsPSObject("CommandLines", out psCmdLines))
                //{
                //    PSObject psCmdLine;
                //    if (psCmdLines.TryGetPropertyAsPSObject("CommandLine", out psCmdLine))
                //    {
                //        example.CommandLines = psCmdLine.GetPropertyValue("commandText");
                //    }
                //    else
                //    {
                //        example.CommandLines = psCmdLines.GetPropertyValue("CommandLine");
                //    }
                //}
                //sb.Clear();

                //PSObject psRemark;
                //if (pso.TryGetPropertyAsPSObject("Remarks", out psRemark))
                //{
                //    if (psRemark.IsArrayList())
                //    {
                //        foreach (PSObject item in psRemark.ToArrayList())
                //        {
                //            sb.AppendLine(item.GetPropertyValue("Text"));
                //        }
                //    }
                //    else
                //        sb.AppendLine(psRemark.GetPropertyValue("Text"));
                //}
                //example.Remarks = sb.ToString();

                _result.Add(example);
            }

            return _result;
        }
        private static List<PSCommandHelpSyntax> GetCommandSyntax(PSObject pso)
        {
            List<PSCommandHelpSyntax> _result = new List<PSCommandHelpSyntax>();
            if (pso.IsArrayList())
            {
                foreach (PSObject syntax in pso.ToArrayList())
                {
                    List<PSCommandHelpSyntax> _r_nested = GetCommandSyntax(syntax);
                    if (_r_nested != null) _result.AddRange(_r_nested);
                }
            }
            else
            {
                PSCommandHelpSyntax syntax = new PSCommandHelpSyntax();
                //syntax.Id = Guid.NewGuid();
                //syntax.CommandHelpId = commandId;
                syntax.Name = pso.GetPropertyValue("Name");
                if (syntax.Name.IsNullOrEmpty()) throw new Exception("Syntax Name can not be null");

                PSObject[] syntaxparams;
                pso.TryGetPropertyAsTObject<PSObject[]>("parameters", out syntaxparams);
                //PSObject p = new PSObject(px);
                //pso.TryGetPropertyAsPSObject("ParametersEx", out parameters);
                syntax.PSCommandHelpSyntaxParameters = new List<PSCommandHelpSyntaxParameter>();
                foreach (var s in syntaxparams)
                {
                    syntax.PSCommandHelpSyntaxParameters.AddRange(GetCommandSyntaxParameters(s));
                }

                //PSObject psParameter;
                //if (pso.TryGetPropertyAsPSObject("parameter", out psParameter))
                //{
                //    syntax.PSCommandHelpSyntaxParameters = GetCommandSyntaxParameters(psParameter);
                //}
                _result.Add(syntax);
            }
            return _result;
        }

        private static List<PSCommandHelpSyntaxParameter> GetCommandSyntaxParameters(PSObject pso)
        {
            List<PSCommandHelpSyntaxParameter> _result = new List<PSCommandHelpSyntaxParameter>();
            if (pso.IsArrayList())
            {
                foreach (PSObject parameter in pso.ToArrayList())
                {
                    List<PSCommandHelpSyntaxParameter> _r_nested = GetCommandSyntaxParameters(parameter);
                    if (_r_nested != null) _result.AddRange(_r_nested);
                }
            }
            else
            {
                PSCommandHelpSyntaxParameter p = new PSCommandHelpSyntaxParameter();
                //p.Id = Guid.NewGuid();
                //p.SyntaxId = syntaxId;
                p.Name = pso.GetPropertyValue("Name");
                if (p.Name.IsNullOrEmpty()) throw new Exception("Syntax Parameter name can not be null/blank");
                p.Required = pso.GetPropertyValue("Required").ParseBool();
                p.ParameterValue = pso.GetPropertyValue("parameterValue");
                p.Position = pso.GetPropertyValue("position");
                p.PipelineInput = pso.GetPropertyValue("PipelineInput").ParseBool();

                //if (p.PipelineInput != null && p.PipelineInput == true)
                //{
                //    p.PipelineInputType = pso.GetPropertyValue("PipelineInputType");

                //    if (!p.PipelineInputType.IsNullOrEmpty())
                //    {
                //        string[] pipelineinputtype = p.PipelineInputType.Split(new string[] { "(", ")" }, StringSplitOptions.RemoveEmptyEntries);
                //        if (pipelineinputtype.Length > 1) p.PipelineInputType = pipelineinputtype[1];
                //    }
                //}
                _result.Add(p);
            }

            return _result;
        }

        private static string GetCommandDetails(PSObject pso)
        {

            StringBuilder sb = new StringBuilder();
            int i, arraylength;
            PSObject psoDescription = (PSObject)pso.Properties["Description"].Value;//).ImmediateBaseObject).Properties["Text"];
            if (psoDescription.ImmediateBaseObject is IList)
            {
                i = 0;
                arraylength = ((ArrayList)psoDescription.ImmediateBaseObject).Count;
                foreach (PSObject item in (ArrayList)psoDescription.ImmediateBaseObject)
                {
                    i++;
                    string counter = string.Empty;
                    counter = arraylength == 1 ? "" : " " + i.ToString();
                    PSPropertyInfo Description = item.Properties["Text"];
                    if (Description != null)
                        sb.AppendLine("Description" + counter + " : " + Description.Value.ToString());
                    else
                        sb.AppendLine("Description" + counter + " : Error in parsing");
                }


            }
            else
            {
                PSPropertyInfo Description = ((PSObject)psoDescription.ImmediateBaseObject).Properties["Text"];
                if (Description != null)
                    sb.AppendLine("Description : " + Description.Value.ToString());
                else
                    sb.AppendLine("Description : Error in parsing");
            }


            //PSPropertyInfo Copyright = ((PSObject)((PSObject)pso.Properties["Copyright"].Value).ImmediateBaseObject).Properties["Text"];
            ////if (Copyright != null)
            //sb.AppendLine("CopyRight : " + Copyright.Value.ToString());

            PSObject psoCopyRight = (PSObject)pso.Properties["Copyright"].Value;
            if (psoCopyRight.ImmediateBaseObject is IList)
            {
                i = 0;
                string counter = string.Empty;
                arraylength = ((ArrayList)psoCopyRight.ImmediateBaseObject).Count;
                counter = arraylength == 1 ? "" : " " + i.ToString();
                foreach (PSObject item in (ArrayList)psoCopyRight.ImmediateBaseObject)
                {
                    i++;
                    PSPropertyInfo CopyRight = item.Properties["Text"];
                    if (CopyRight != null)
                        sb.AppendLine("CopyRight" + counter + " : " + CopyRight.Value.ToString());
                    else
                        sb.AppendLine("CopyRight" + counter + " : Error in parsing");
                }

            }
            else
            {
                PSPropertyInfo CopyRight = ((PSObject)psoCopyRight.ImmediateBaseObject).Properties["Text"];
                if (CopyRight != null)
                    sb.AppendLine("CopyRight : " + CopyRight.Value.ToString());
                else
                    sb.AppendLine("CopyRight : Error in parsing");
            }


            sb.AppendLine("Noun : " + pso.Properties["Noun"].Value.ToString());
            sb.AppendLine("Verb : " + pso.Properties["Verb"].Value.ToString());
            sb.AppendLine("Name : " + pso.Properties["Name"].Value.ToString());
            sb.AppendLine("Version : " + pso.Properties["Version"].Value.ToString());

            return sb.ToString();
        }
        private static string GetCommandReturnValue(PSObject pso)
        {

            StringBuilder sb = new StringBuilder();

            PSPropertyInfo piReturnValue = pso.Properties["ReturnValue"];
            if (piReturnValue != null && piReturnValue.Value != null && piReturnValue.Value is PSObject)
            {
                PSObject psReturnValue = (PSObject)piReturnValue.Value;

                sb.Clear();
                List<string> types = new List<string>();
                List<string> descriptions = new List<string>();
                if (psReturnValue.ImmediateBaseObject is IList)
                {
                    int i;//, count;
                    i = 0;

                    foreach (PSObject psreturnvalue in (ArrayList)psReturnValue.ImmediateBaseObject)
                    {
                        i++;

                        PSPropertyInfo piType = psreturnvalue.Properties["Type"];
                        if (piType != null && piType.Value != null && piType.Value is PSObject)
                        {
                            PSObject psType = (PSObject)piType.Value;
                            if (psType.ImmediateBaseObject is PSObject)
                            {

                                PSPropertyInfo piName = ((PSObject)psType.ImmediateBaseObject).Properties["Name"];
                                if (piName != null && piName.Value != null && piName.Value is PSObject)
                                {
                                    PSObject Name = (PSObject)piName.Value;

                                    if (Name.ImmediateBaseObject is IList)
                                        foreach (PSObject item in (ArrayList)Name.ImmediateBaseObject)
                                        {
                                            PSPropertyInfo typeText = item.Properties["Text"];
                                            if (typeText != null) types.Add("Type" + i.ToString("000") + " : " + typeText.Value.ToString());
                                        }
                                    else
                                    {
                                        PSPropertyInfo typeText = Name.Properties["Text"];
                                        if (typeText != null) types.Add("Type : " + typeText.Value.ToString());
                                    }

                                }

                                PSPropertyInfo piDescription = psreturnvalue.Properties["Description"];
                                if (piDescription != null && piDescription.Value != null && piDescription.Value is PSObject)
                                {
                                    PSObject Description = (PSObject)piDescription.Value;
                                    if (Description.ImmediateBaseObject is IList)
                                        foreach (PSObject item in (ArrayList)Description.ImmediateBaseObject)
                                        {
                                            PSPropertyInfo Text = item.Properties["Text"];
                                            if (Text != null && Text.Value != null) descriptions.Add("Description" + i.ToString("000") + " : " + Text.Value.ToString());
                                        }
                                    else
                                    {
                                        PSPropertyInfo Text = Description.Properties["Text"];
                                        if (Text != null && Text.Value != null) descriptions.Add("Description : " + Text.Value.ToString());
                                    }
                                }
                            }
                        }
                    }
                    i = 0;
                    foreach (string item in types)
                    {
                        sb.AppendLine(item);
                        if (i < descriptions.Count) sb.AppendLine(descriptions[i]);
                        i++;
                    }

                }
                else
                {
                    PSObject psType = (PSObject)psReturnValue.Properties["Type"].Value;
                    PSPropertyInfo piName = ((PSObject)psType.ImmediateBaseObject).Properties["Name"];
                    if (piName != null)
                    {
                        if (piName.Value is PSObject)
                        {
                            PSObject Name = (PSObject)piName.Value;

                            if (Name.ImmediateBaseObject is IList)
                                foreach (PSObject item in (ArrayList)Name.ImmediateBaseObject)
                                {
                                    PSPropertyInfo typeText = item.Properties["Text"];
                                    if (typeText != null) sb.AppendLine("Type : " + typeText.Value.ToString());
                                }
                            else
                                if (Name.Properties["Text"] != null) sb.AppendLine("Type : " + Name.Properties["Text"].Value.ToString());
                        }
                        else
                            sb.AppendLine("Type : " + piName.Value.ToString());
                    }


                    PSObject Description = (PSObject)psReturnValue.Properties["Description"].Value;
                    if (Description.ImmediateBaseObject is IList)
                        foreach (PSObject item in (ArrayList)Description.ImmediateBaseObject)
                        {
                            PSPropertyInfo Text = item.Properties["Text"];
                            if (Text != null) sb.AppendLine("Description : " + Text.Value.ToString());
                        }

                }
            }
            return sb.ToString();
        }

        #endregion

        private Runspace GetExecutionRunspace()
        {
            //if (!Ispersistant)
            //    CreateRunspace();
            Runspace runspace = null;
            bool skipSettingDefaultRunspace = false;
            switch (RunspaceMode)
            {
                case PowerShellRunspaceMode.RemoteRunspace:
                    if (_remoteSessionLocalRunspaceForDefultRunspace == null || _remoteSessionLocalRunspaceForDefultRunspace.RunspaceStateInfo.State != RunspaceState.Opened)
                    {
                        _remoteSessionLocalRunspaceForDefultRunspace = CreateLocalDefaultRunspace(ImportedModuleList);
                        SetupdefinedVariables(_remoteSessionLocalRunspaceForDefultRunspace);
                    }
                    System.Management.Automation.Runspaces.Runspace.DefaultRunspace = _remoteSessionLocalRunspaceForDefultRunspace;
                    skipSettingDefaultRunspace = true;
                    if (Runspace.RunspaceStateInfo.State != RunspaceState.Opened)
                        Create();

                    runspace = this.Runspace;
                    break;
                case PowerShellRunspaceMode.RemoteSessionImported:
                case PowerShellRunspaceMode.LocalRunspace:
                    if (Runspace.RunspaceStateInfo.State != RunspaceState.Opened)
                        Create();

                    runspace = this.Runspace;
                    break;
                case PowerShellRunspaceMode.RemoteSession:
                    if (ExecutionRunspace.PSSession.Runspace.RunspaceStateInfo.State != RunspaceState.Opened)
                        Create();
                    if (_remoteSessionLocalRunspaceForDefultRunspace == null || _remoteSessionLocalRunspaceForDefultRunspace.RunspaceStateInfo.State != RunspaceState.Opened)
                    {
                        _remoteSessionLocalRunspaceForDefultRunspace = CreateLocalDefaultRunspace(ImportedModuleList);
                        SetupdefinedVariables(_remoteSessionLocalRunspaceForDefultRunspace);
                        //System.Management.Automation.Runspaces.Runspace.DefaultRunspace = runspace;
                    }
                    runspace = _remoteSessionLocalRunspaceForDefultRunspace;
                    break;
                default:
                    throw new NotImplementedException();
            }

            //set default runspace for thread. DefaultRunspace is required pre thread in multithreading or async operation.
            //if (System.Management.Automation.Runspaces.Runspace.DefaultRunspace == null)
            if (!skipSettingDefaultRunspace)
                System.Management.Automation.Runspaces.Runspace.DefaultRunspace = runspace;

            return runspace;
        }

        private void CreateRunspace(Action<PSDataStreams> psDataStreamAction)
        {
            InitialSessionState iss = InitialSessionState.CreateDefault();
            switch (RunspaceMode)
            {
                case PowerShellRunspaceMode.RemoteRunspace:
                    PSCredential credential = new PSCredential(UserName, Password);
                    WSManConnectionInfo connectionInfo = GetWsManConnection();
                    connectionInfo.MaximumConnectionRedirectionCount = MaxRedirectionCount;
                    this.Runspace = RunspaceFactory.CreateRunspace(connectionInfo);
                    this.Runspace.Open();
                    //System.Management.Automation.Runspaces.Runspace.DefaultRunspace = this.Runspace;
                    RunspaceCreated = true;

                    AttachRunspaceChanged(this.Runspace);
                    break;

                case PowerShellRunspaceMode.RemoteSession:
                    ExecutionRunspace = new ExecutionRunspace(UserName, Password, ConnectionUri, SchemaUri, MaxRedirectionCount > 0, false, ImportedModuleList, psDataStreamAction);
                    this.Runspace = ExecutionRunspace.Runspace;
                    RunspaceCreated = true;

                    if (ExecutionRunspace.PSSession != null)
                        AttachRunspaceChanged(ExecutionRunspace.PSSession.Runspace);
                    break;

                case PowerShellRunspaceMode.RemoteSessionImported:
                    ExecutionRunspace = new ExecutionRunspace(UserName, Password, ConnectionUri, SchemaUri, MaxRedirectionCount > 0, true, ImportedModuleList, psDataStreamAction);
                    this.Runspace = ExecutionRunspace.Runspace;
                    RunspaceCreated = true;

                    if (ExecutionRunspace.PSSession != null)
                        AttachRunspaceChanged(ExecutionRunspace.PSSession.Runspace);
                    break;

                case PowerShellRunspaceMode.LocalRunspace:
                    //if (ImportedModuleList == null || ImportedModuleList.Count == 0) ImportedModuleList = new List<string>() { "Microsoft.PowerShell.*" };
                    ExecutionRunspace = new ExecutionRunspace(ImportedModuleList, psDataStreamAction: psDataStreamAction);
                    this.Runspace = ExecutionRunspace.Runspace;
                    RunspaceCreated = true;

                    AttachRunspaceChanged(this.Runspace);
                    break;

                default:
                    throw new NotImplementedException("Unknown or not implemented RunspaceMode : " + RunspaceMode);
            }

            //SetupdefinedVariables(GetExecutionRunspace());
            GetExecutionRunspace();
        }


        private static Collection<PSCommandInfoProxy> ConvertScriptedCommandInfoToProxy(Collection<PSObject> psObjectCollection, bool includeCommonParameters = false)
        {
            Collection<PSCommandInfoProxy> _result = new Collection<PSCommandInfoProxy>();
            foreach (var psobj in psObjectCollection)
            {
                PSCommandInfoProxy _colItem = new PSCommandInfoProxy();
                dynamic commandInfoDynamic = psobj;
                CommandTypes commandType = Enum.Parse(typeof(CommandTypes), commandInfoDynamic.CommandType);
                _colItem.IsRemoteCommand = false;
                _colItem.Name = commandInfoDynamic.Name;
                _colItem.NameLower = _colItem.Name.ToLower();
                _colItem.CommandType = commandType.ToString();
                _colItem.Definition = commandInfoDynamic.Definition.Trim();
                _colItem.Module = commandInfoDynamic.Module;
                _colItem.ModuleName = commandInfoDynamic.ModuleName;

                _colItem.OutputType = commandInfoDynamic.OutputType;
                _colItem.ParameterSets = commandInfoDynamic.ParameterSets;

                //?++ fix includecommon parameters
                _colItem.Parameters = commandInfoDynamic.Parameters;
                //_res.Parameters = _res.Parameters.Where
                //if (includeCommonParameters)
                //    _res.Parameters = (from a in commandInfoDynamic.Parameters select a.Key).ToArray();
                //else
                //    _res.Parameters = (from a in commandInfoDynamic.Parameters where !Constants.CommomParamaters.Contains(a.Key) select a.Key).ToArray();

                _colItem.RemotingCapability = commandInfoDynamic.RemotingCapability;

                _colItem.Visibility = commandInfoDynamic.Visibility;

                _colItem.Noun = commandInfoDynamic.Noun;
                _colItem.Options = commandInfoDynamic.Options;
                _colItem.Description = commandInfoDynamic.Description;
                _colItem.ReferencedCommand = commandInfoDynamic.ReferencedCommand;
                _colItem.ResolvedCommand = commandInfoDynamic.ResolvedCommand;

                _colItem.DefaultParameterSet = commandInfoDynamic.DefaultParameterSet;
                _colItem.HelpFile = commandInfoDynamic.HelpFile;
                //_res.HelpUri = commandInfoDynamic.HelpUri;
                _colItem.Verb = commandInfoDynamic.Verb;
                _colItem.ImplementingType = commandInfoDynamic.ImplementingType;
                _colItem.PSSnapIn = commandInfoDynamic.PSSnapIn;
                _colItem.OriginalEncoding = commandInfoDynamic.OriginalEncoding;
                _colItem.Path = commandInfoDynamic.Path;
                _colItem.ScriptBlock = commandInfoDynamic.ScriptBlock;
                _colItem.ScriptContents = commandInfoDynamic.ScriptContents;
                _colItem.CmdletBinding = commandInfoDynamic.CmdletBinding;

                _result.Add(_colItem);
                //switch (commandType)
                //{
                //    case CommandTypes.Alias:
                //        _res.Noun = commandInfoDynamic.Noun;
                //        if (commandInfoDynamic.Options != null) _res.Options = commandInfoDynamic.Options.ToString();
                //        _res.Description = commandInfoDynamic.Description;
                //        if (commandInfoDynamic.ReferencedCommand != null) _res.ReferencedCommand = commandInfoDynamic.ReferencedCommand.ToString();
                //        if (commandInfoDynamic.ResolvedCommand != null) _res.ResolvedCommand = commandInfoDynamic.ResolvedCommand.ToString();
                //        break;
                //    case CommandTypes.Cmdlet:
                //        _res.Noun = commandInfoDynamic.Noun;
                //        if (commandInfoDynamic.Options != null) _res.Options = commandInfoDynamic.Options.ToString();
                //        _res.DefaultParameterSet = commandInfoDynamic.DefaultParameterSet;
                //        _res.HelpFile = commandInfoDynamic.HelpFile;
                //        if (commandInfoDynamic.HelpUri != null) _res.HelpUri = commandInfoDynamic.HelpUri.ToString();
                //        _res.Verb = commandInfoDynamic.Verb;
                //        if (commandInfoDynamic.ImplementingType != null) _res.ImplementingType = commandInfoDynamic.ImplementingType.FullName;
                //        if (commandInfoDynamic.PSSnapIn != null) _res.PSSnapIn = commandInfoDynamic.PSSnapIn.Name;
                //        break;
                //    case CommandTypes.ExternalScript:
                //        if (commandInfoDynamic.OriginalEncoding != null) _res.OriginalEncoding = commandInfoDynamic.OriginalEncoding.ToString();
                //        _res.Path = commandInfoDynamic.Path;
                //        if (commandInfoDynamic.ScriptBlock != null) _res.ScriptBlock = commandInfoDynamic.ScriptBlock.ToString();
                //        _res.ScriptContents = commandInfoDynamic.ScriptContents;
                //        break;
                //    case CommandTypes.Function:
                //        _res.Noun = commandInfoDynamic.Noun;
                //        if (commandInfoDynamic.Options != null) _res.Options = commandInfoDynamic.Options.ToString();
                //        _res.DefaultParameterSet = commandInfoDynamic.DefaultParameterSet;
                //        _res.HelpFile = commandInfoDynamic.HelpFile;
                //        if (commandInfoDynamic.HelpUri != null) _res.HelpUri = commandInfoDynamic.HelpUri.ToString();
                //        _res.Verb = commandInfoDynamic.Verb;

                //        _res.CmdletBinding = commandInfoDynamic.CmdletBinding;
                //        _res.Description = commandInfoDynamic.Description;
                //        if (commandInfoDynamic.ScriptBlock != null) _res.ScriptBlock = commandInfoDynamic.ScriptBlock.ToString();
                //        break;
                //    case CommandTypes.Script:
                //        if (commandInfoDynamic.ScriptBlock != null) _res.ScriptBlock = commandInfoDynamic.ScriptBlock.ToString();
                //        break;
                //    default:
                //        break;
                //}

            }
            return _result;
        }

        /// <summary>
        /// Runs the given powershell script and returns the script output.
        /// </summary>
        /// <param name="scriptText">the powershell script text to run</param>
        /// <returns>output of the script</returns> c
        //private string GetScriptOutput(string scriptText)
        //{
        //    ResetRunspace();
        //    StringBuilder sb = new StringBuilder();

        //    using (PowerShell powerShell = PowerShell.Create())
        //    {
        //        if (RemoteRunSpace.RunspaceStateInfo.State != RunspaceState.Opened) RemoteRunSpace.Open();

        //        //Pipeline Pipeline = RemoteRunspace.CreatePipeline();
        //        //Pipeline.Commands.AddScript(scriptText);

        //        //// add an extra command to transform the script output objects into nicely formatted strings
        //        //// remove this line to get the actual objects that the script returns. For example, the script
        //        //// "Get-Process" returns a collection of System.Diagnostics.Process instances.
        //        //Pipeline.Commands.Add("Out-String");

        //        powerShell.Runspace = RemoteRunSpace;

        //        powerShell.AddScript(scriptText + " | out-string");
        //        //powerShell.AddCommand("Out-String");

        //        Collection<PSObject> _result;
        //        try
        //        {
        //            //_result = Pipeline.Invoke();
        //            _result = powerShell.Invoke();
        //        }
        //        catch (Exception ex)
        //        {
        //            throw ex;
        //        }

        //        // convert the script result into a single string

        //        foreach (PSObject obj in _result)
        //        {
        //            sb.AppendLine(obj.ToString());
        //        }
        //    }


        //    return sb.ToString();
        //}
        //private string GetScriptOutput(string scriptText, Runspace runSpace)
        //{

        //    Pipeline Pipeline = runSpace.CreatePipeline();
        //    Pipeline.Commands.AddScript(scriptText);
        //    Pipeline.Commands.Add("Out-String");

        //    Collection<PSObject> _result;

        //    try
        //    {
        //        _result = Pipeline.Invoke();
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }

        //    // convert the script result into a single string
        //    StringBuilder sb = new StringBuilder();
        //    foreach (PSObject obj in _result)
        //    {
        //        sb.AppendLine(obj.ToString());
        //    }

        //    return sb.ToString();
        //}

        /// <summary>
        /// Creates and open a local runspace with default initial session state and list of imported modules.
        /// </summary>
        /// <param name="importedModuleList"></param>
        /// <returns></returns>
        private static Runspace CreateLocalDefaultRunspace(List<string> importedModuleList = null)
        {
            InitialSessionState iss = InitialSessionState.CreateDefault();
            if (importedModuleList != null && importedModuleList.Count > 1) iss.ImportPSModule(importedModuleList.ToArray());
            var runspace = RunspaceFactory.CreateRunspace(iss);
            runspace.Open();
            return runspace;
        }

    }





}
