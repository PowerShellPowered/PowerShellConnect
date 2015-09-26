using System;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;

namespace PowerShellPowered.PowerShellConnect
{
    internal static class iExecutionMethods
    {
        /// <summary>
        /// Execute a powershell command.
        /// </summary>
        /// <param name="command">Powershell command.</param>
        /// <returns>Collection of PSObjects</returns>
        public static Collection<PSObject> ExecuteCommand(PSCommand command, Runspace runspace)
        {
            return ExecuteCommand<PSObject>(command, runspace);
        }

        /// <summary>
        /// Execute a powershell command.
        /// </summary>
        /// <typeparam name="T">Type of command execution result.</typeparam>
        /// <param name="command">Powershell command.</param>
        /// <returns>Result collection.</returns>
        public static Collection<T> ExecuteCommand<T>(PSCommand command, Runspace runspace)
        {
            if (command == null)
            {
                throw new ArgumentOutOfRangeException("command");
            }

            using (PowerShell ps = PowerShell.Create())
            {
                ps.Commands = command;
                ps.Runspace = runspace;

                return ps.Invoke<T>();
            }
        }

        /// <summary>
        /// Method to create PS credential and set it as PS environment variable.
        /// </summary>
        /// <param name="runspace">Runspace to run the command.</param>
        /// <param name="userName">The user name.</param>
        /// <param name="password">The user password.</param>
        public static void SetCredential(Runspace runspace, string userName, SecureString password)
        {
            if (runspace == null)
            {
                throw new ArgumentOutOfRangeException("runspace");
            }

            if (string.IsNullOrEmpty(userName))
            {
                throw new ArgumentOutOfRangeException("userName");
            }

            if (password == null)
            {
                throw new ArgumentOutOfRangeException("password");
            }

            var cred = new PSCredential(userName, password);
            SetVariable(runspace, Constants.PSVariableNameStrings.Credential, cred);
        }

        /// <summary>
        /// Method to set passed in object to PS environment variable
        /// </summary>
        /// <param name="runspace">Runspace to run the command.</param>
        /// <param name="name">name of the variable</param>
        /// <param name="value">value of the variable</param>
        public static void SetVariable(Runspace runspace, string name, object value)
        {
            if (runspace == null)
            {
                throw new ArgumentOutOfRangeException("runspace");
            }

            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentOutOfRangeException("name");
            }

            if (value == null)
            {
                throw new ArgumentOutOfRangeException("value");
            }

            var command = new PSCommand();
            command.AddCommand(Constants.PSCmdlets.SetVariable);
            command.AddParameter(Constants.PSParameterNameStrings.Name, name);
            command.AddParameter(Constants.PSParameterNameStrings.Value, value);
            ExecuteCommand<PSObject>(command, runspace);
        }

        public static string FormatPowerShellCommandName(string name)
        {
            string[] cmd = name.Split(new string[] { "-" }, StringSplitOptions.RemoveEmptyEntries);
            cmd[0] = cmd[0].Substring(0, 1).ToUpper() + cmd[0].Substring(1).ToLower();
            cmd[1] = cmd[1].Substring(0, 1).ToUpper() + cmd[1].Substring(1).ToLower();
            return string.Join("-", cmd);
        }
    }
}