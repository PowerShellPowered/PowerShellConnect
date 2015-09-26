using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;

namespace PowerShellPowered.PowerShellConnect
{
    internal class ExecutionRunspace : IDisposable
    {
        ExecutionSession executionSession;

        public Runspace Runspace { get; protected set; }

        public PSSession PSSession { get; protected set; }

        public bool AllowRedirection { get; set; }

        public bool SessionImported { get; set; }
        public PSModuleInfo SessionModule { get; set; }

        public ExecutionRunspace(List<string> importedModuleList = null, Action<PSDataStreams> psDataStreamAction = null)
        {
            InitialSessionState iss = InitialSessionState.CreateDefault2();
            if (importedModuleList != null && importedModuleList.Count > 0) iss.ImportPSModule(importedModuleList.ToArray());
            this.Runspace = RunspaceFactory.CreateRunspace(iss);
            this.Runspace.Open();
            System.Management.Automation.Runspaces.Runspace.DefaultRunspace = this.Runspace;

        }

        //?+ TODO: create another to use computername
        //? no need to create new, use https://computer:
        /// <summary>
        /// Create Connection to remote server using bacis credential.        
        /// </summary>
        /// <param name="userName">user name</param>
        /// <param name="password">password in secure string format</param>
        /// <param name="connectionUri">
        /// Specifies a Uniform Resource Identifier (URI) that defines the connection endpoint for the session. The URI must be fully qualified. The format of this string is as follows:
        ///<![CDATA[
        ///<Transport://<ComputerName>:<Port>/<ApplicationName>
        ///he default value is as follows
        ///http://localhost:5985/WSMAN 
        ///]]>
        ///If you do not specify a ConnectionURI, you can use the UseSSL, ComputerName, Port, and ApplicationName parameters to specify the ConnectionURI values.
        ///Valid values for the Transport segment of the URI are HTTP and HTTPS. If you specify a connection URI with a Transport segment, but do not specify a port, the session is created with standards ports: 80 for HTTP and 443 for HTTPS. To use the default ports for Windows PowerShell remoting, specify port 5985 for HTTP or 5986 for HTTPS.
        ///</param>
        /// <param name="schemaUri"> schema or configuration name</param>
        /// <param name="allowRedirection"> allow redirection of session</param>
        /// <param name="importSession"> import the session</param>
        /// <param name="connectionInfo">WSManConnection infor object</param>
        public ExecutionRunspace(string userName, SecureString password, string connectionUri, string schemaUri,
                bool allowRedirection = false, bool importSession = false, List<string> importedModuleList = null, Action<PSDataStreams> psDataStreamAction = null)
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

            AllowRedirection = allowRedirection;

            InitialSessionState iss = InitialSessionState.CreateDefault();
            if (importedModuleList != null && importedModuleList.Count > 1) iss.ImportPSModule(importedModuleList.ToArray());
            this.Runspace = RunspaceFactory.CreateRunspace(iss);
            this.Runspace.Open();
            System.Management.Automation.Runspaces.Runspace.DefaultRunspace = this.Runspace;
            //Create the session and import it.
            executionSession = new ExecutionSession(this.Runspace, AllowRedirection);
            this.PSSession = executionSession.Create(userName, password, connectionUri, schemaUri, psDataStreamAction);

            if (this.PSSession == null) throw new InvalidOperationException("can not connect to PS Session");

            if (importSession)
                ImportSession(psDataStreamAction);
        }

        internal void ImportSession(Action<PSDataStreams> psDataStreamAction)
        {
            if (SessionImported) return;

            if (executionSession != null && this.PSSession != null)
            {
                SessionModule = executionSession.ImportPSSession(this.PSSession, psDataStreamAction);
                SessionImported = true;
            }
            else throw new PSInvalidOperationException("Can not import session, the Session does not exist");

        }

        public Collection<PSObject> ExecuteCommand(PSCommand command, Action<PSDataStreams> psDataStreamAction = null)
        {
            return this.Runspace.ExecuteCommand<PSObject>(command, psDataStreamAction);
            //return this.ExecuteCommand<PSObject>(command, psDataStreamAction);
        }

        public Collection<T> ExecuteCommand<T>(PSCommand command, Action<PSDataStreams> psDataStreamAction = null)
        {
            return this.Runspace.ExecuteCommand<T>(command, psDataStreamAction);
            //if (command == null)
            //{
            //    throw new ArgumentOutOfRangeException("command");
            //}

            //using (PowerShell ps = PowerShell.Create())
            //{
            //    ps.Commands = command;
            //    ps.Runspace = this.Runspace;
            //    var result = ps.Invoke<T>();
            //    if (psDataStreamAction != null) psDataStreamAction(ps.Streams);

            //    return result;

            //}
        }

        public Collection<PSObject> ExecuteScript(string scriptText, Dictionary<string, object> scriptParameters = null, bool out_String = false, bool stream = false, int width = 10000, Action<PSDataStreams> psDataStreamAction = null)
        {
            return this.Runspace.ExecuteScript<PSObject>(scriptText, scriptParameters, out_String, stream, width, psDataStreamAction);

            //return this.ExecutePipeLine<PSObject>(scriptText, out_String, stream, width, psDataStreamAction);
        }

        public Collection<T> ExecuteScript<T>(string scriptText, Dictionary<string, object> scriptParameters = null, bool out_String = false, bool stream = false, int width = 10000, Action<PSDataStreams> psDataStreamAction = null)
        {
            //?++ using runspace extension method to simplify management of code.
            return this.Runspace.ExecuteScript<T>(scriptText, scriptParameters, out_String, stream, width, psDataStreamAction);

            //if (scriptText == null)
            //{
            //    throw new ArgumentOutOfRangeException("scriptText");
            //}

            //using (PowerShell ps = PowerShell.Create())
            //{
            //    ps.Runspace = this.Runspace;
            //    ps.AddScript(scriptText);
            //    if (out_String && typeof(T).Equals(typeof(string)))
            //    {
            //        ps.AddCommand("Out-String").AddParameter("Width", width);

            //        if (stream) ps.AddParameter("Stream");                    

            //    }
            //    //ps.AddCommand("Out-String");
            //    if(psDataStreamAction != null) 
            //        psDataStreamAction(ps.Streams);
            //    var result = ps.Invoke<T>();                
            //    return result;
            //}
        }

        //[Obsolete]
        //public string ExecutePipeLine(string scriptText, string v, bool stream = false, int width = 10000, Action<ProgressRecord> callback = null)
        //{
        //    if (scriptText == null)
        //    {
        //        throw new ArgumentOutOfRangeException("scriptText");
        //    }

        //    StringBuilder sb = new StringBuilder();
        //    using (PowerShell ps = PowerShell.Create())
        //    {
        //        ps.Runspace = this.Runspace;
        //        ps.AddScript(scriptText);
        //        ps.AddCommand("Out-String");
        //        if (stream) ps.AddParameter("Stream");
        //        ps.AddParameter("Width", width);
        //        //ps.Streams.Progress.DataAdded += Progress_DataAdded;
        //        ps.Streams.Progress.DataAdded += (x, y) =>
        //        {
        //            var vm = x as PSDataCollection<ProgressRecord>;
        //            if (vm != null)
        //            {
        //                callback(vm[y.Index]);
        //            }
        //        };
        //        try
        //        {
        //            //var vvs = ps.Invoke<string>();
        //        }
        //        catch (Exception)
        //        {

        //            //throw;
        //        }
        //        Collection<PSObject> _result = ps.Invoke();

        //        foreach (var item in _result)
        //        {
        //            string value = item.ToString();// +" -------" + item.ToString().Length;
        //            if (stream) sb.AppendLine(value);
        //            else sb.Append(value);
        //        }

        //    }
        //    return sb.ToString();
        //}

        internal void CloseInstance(Action<PSDataStreams> psDataStreamAction = null)
        {
            try
            {
                //var remotingSession = new ExecutionSession(this.Runspace);
                //remotingSession.RemovePSSession(this.PSSession);

                if (this.PSSession != null)
                    executionSession.RemovePSSession(this.PSSession, psDataStreamAction);

                this.Runspace.Close();
                SessionImported = false;
                //this.Runspace = null;
            }
            catch
            {
            }
        }

        ~ExecutionRunspace()
        {
            CloseInstance();
        }

        public void Dispose()
        {
            //CloseInstance();
        }
    }
}