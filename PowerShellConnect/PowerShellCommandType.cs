
namespace PowerShellPowered.PowerShellConnect
{
    public enum PowerShellCommandTypes
    {
        Alias = 1,
        Function = 2,
        Filter = 4,
        Cmdlet = 8,
        ExternalScript = 16,
        Application = 32,
        Script = 64,
        All = 127,
        AllPowerShellNative = 90,
    }
}
