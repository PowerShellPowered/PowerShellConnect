
namespace PowerShellPowered.PowerShellConnect
{
    /// <summary>
    /// Creates shell based on Powershell configuration types
    /// </summary>
    public enum ShellRunspaceMode
    {
        /// <summary>
        /// Local RunSpace in PowerShell
        /// </summary>
        Local = 0,
        /// <summary>
        /// Remote RunSpace
        /// </summary>
        Remote = 1,
        /// <summary>
        /// Remote Session by creating PSSession object.\r\nPSSession is not imported in RunSpace
        /// </summary>
        RemoteSession = 2,
        /// <summary>
        /// Remote Session by creating PSSession object and Importing it in RunSpace. Scripts can use Remote available commands. 
        /// </summary>
        RemoteSessionImported = 3,
    }
}
