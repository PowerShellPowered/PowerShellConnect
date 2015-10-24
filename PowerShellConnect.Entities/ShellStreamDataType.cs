using System;

namespace PowerShellPowered.Entities
{
    [Flags]
    public enum ShellStreamDataType
    {
        None = 0,
        Debug = 1,
        Error = 2,
        GenericError = 4,
        Progress = 8,
        Verbose = 16,
        Warning = 32,

        All = 63,
    }

}
