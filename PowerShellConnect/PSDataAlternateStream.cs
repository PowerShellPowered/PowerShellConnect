using System.Management.Automation;

namespace PowerShellPowered.PowerShellConnect
{
    public class PSDataAlternateStream
    {
        public string SimpleError { get; set; }
        public ErrorRecord Error { get; set; }
        public ProgressRecord SimpleProgress { get; set; }

        public bool HasError { get; set; }
        public bool HasProgress { get; set; }
        public bool HasSimpleError { get; set; }
    }
}
