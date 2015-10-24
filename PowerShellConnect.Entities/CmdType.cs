
namespace PowerShellPowered.Entities
{
    public enum CmdType
    {
        Alias = 1,
        Function = 2,
        Filter = 4,
        Cmdlet = 8,
        ExternalScript = 16,
        Application = 32,
        Script = 64,
        Workflow = 128,
        Configuration = 256,
        All = 511,
        AllPowerShellNative = 90,
    }
}
