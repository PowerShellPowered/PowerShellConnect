using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace PowerShellPowered.PowerShellConnect.Entities
{
    public class CmdInfo
    {
        public bool IsRemoteCommand { get; set; }
        public bool CmdletBinding { get; set; }
        public CmdType CommandType { get; set; }
        public string DefaultParameterSet { get; set; }
        public string Definition { get; set; }
        public string Description { get; set; }
        public string HelpFile { get; set; }
        public string HelpUri { get; set; }
        public string ImplementingType { get; set; }
        public string Module { get; set; }
        public string ModuleName { get; set; }
        public string Name { get; set; }
        public string NameLower { get; set; }
        public string Noun { get; set; }
        public string Options { get; set; }
        public string OriginalEncoding { get; set; }
        public List<string> OutputType { get; set; } = new List<string>();
        public List<string> Parameters { get; set; } = new List<string>();
        public List<string> ParameterSets { get; set; } = new List<string>();
        public List<CmdParameterInfo> CmdParameters { get; set; } = new List<CmdParameterInfo>();
        public List<CmdParameterSetInfo> CmdParameterSets { get; set; } = new List<CmdParameterSetInfo>();
        public string Path { get; set; }
        public string PSSnapIn { get; set; }
        public string ReferencedCommand { get; set; }
        public string RemotingCapability { get; set; }
        public string ResolvedCommand { get; set; }
        public string ScriptBlock { get; set; }
        public string ScriptContents { get; set; }
        public string Verb { get; set; }
        public string Visibility { get; set; }

    }

    public class CmdParameterInfo
    {
        public string Name { get; set; }
        public string NameLower { get; set; }
        public string ParameterType { get; set; }
        public bool IsMandatory { get; set; }
        public bool IsDynamic { get; set; }
        public int Position { get; set; }
        public bool ValueFromPipeline { get; set; }
        public List<string> ValidateSetValues { get; set; } = new List<string>();
        //public bool ValueFromPipelineByPropertyName { get; set; }
        //public bool ValueFromRemainingArguments { get; set; }
        //public string HelpMessage { get; set; }
        //public List<string> Aliases { get; set; }        
        public bool SwitchParameter { get; set; }
    }

    public class CmdParameterSetInfo
    {        
        public string Name { get; set; }
        public string NameLower { get; set; }
        public bool IsDefault { get; set; }
        public string ToStringValue { get; set; }
        public List<CmdParameterInfo> Parameters { get; set; } = new List<CmdParameterInfo>();

        public override string ToString()
        {
            return ToStringValue;
        }
    }

    public class ModuleInfo
    {
        public Guid Guid { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public List<string> NestedModules { get; set; }
        public List<string> RequiredModules { get; set; }
        public string ModuleType { get; set; }
        public string RootModule { get; set; }
        public string CompanyName { get; set; }
        public string Author { get; set; }
        public string HelpInfoUri { get; set; }
    }
}
