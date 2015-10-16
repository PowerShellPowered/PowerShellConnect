using System;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;


namespace PowerShellPowered.PowerShellConnect
{
    internal class ExecutionSession
    {
        //?+ TODO: Add option for multiple authentication type, this is just basic at this time.
        private readonly Runspace runspace;
        public bool AllowRedirection { get; set; }
        public ExecutionSession(Runspace runspace, bool allowRedirection = false)
        {
            if (runspace == null)
            {
                throw new ArgumentOutOfRangeException("runspace");
            }

            this.runspace = runspace;
            AllowRedirection = allowRedirection;
        }

        public PSSession Create(string userName, SecureString password, string connectionUri, string schemauri, Action<PSDataStreams> psDataStreamAction)
        {
            if (string.IsNullOrEmpty(userName))
            {
                throw new ArgumentOutOfRangeException("userName");
            }

            if (password == null)
            {
                throw new ArgumentOutOfRangeException("password");
            }

            if (string.IsNullOrEmpty(connectionUri))
            {
                throw new ArgumentOutOfRangeException("connectionUri");
            }
            this.runspace.SetCredentialVariable(userName, password);
            var command = new PSCommand();
            string importpssessionscript = Constants.SessionScripts.NewPSSessionScriptWithBasicAuth;
            if (AllowRedirection) importpssessionscript += Constants.SessionScripts.AllowRedirectionInNewPSSession;
            command.AddScript(string.Format(importpssessionscript, schemauri, connectionUri));
            Collection<PSSession> sessions = this.runspace.ExecuteCommand<PSSession>(command, psDataStreamAction);
            if (sessions.Count > 0) this.runspace.SetRunspaceVariable(Constants.ParameterNameStrings.Session, sessions[0]);
            return sessions.Count == 0 ? null : sessions[0];
        }

        public PSModuleInfo ImportPSSession(PSSession session, Action<PSDataStreams> psDataStreamAction)
        {
            if (session == null)
            {
                throw new ArgumentOutOfRangeException("session");
            }

            var command = new PSCommand();
            command.AddCommand(Constants.SessionScripts.ImportPSSession);
            command.AddParameter(Constants.ParameterNameStrings.Session, session);
            Collection<PSModuleInfo> modules;
            try
            {
                modules = this.runspace.ExecuteCommand<PSModuleInfo>(command, psDataStreamAction);
                if (modules.Count > 0) return modules[0];
            }
            catch (Exception)
            {

                return null;
            }
            return null;
        }

        public void RemovePSSession(PSSession session, Action<PSDataStreams> psDataStreamAction)
        {
            if (session == null)
            {
                throw new ArgumentOutOfRangeException("session");
            }

            var command = new PSCommand();
            command.AddCommand(Constants.SessionScripts.RemovePSSession);
            command.AddParameter(Constants.ParameterNameStrings.Session, session);
            this.runspace.ExecuteCommand(command, psDataStreamAction);
        }


    }
}
