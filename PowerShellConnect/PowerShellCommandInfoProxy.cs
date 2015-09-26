using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace PowerShellPowered.PowerShellConnect.Entities
{
    public class PSCommandInfoProxy
    {
        public bool IsRemoteCommand { get; set; }
        public bool CmdletBinding { get; set; }
        public string CommandType { get; set; }
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
        public string[] OutputType { get; set; }
        public string[] Parameters { get; set; }
        public Collection<PSParameterInfoProxy> ParameterCollection { get; set; }
        public string[] ParameterSets { get; set; }
        public List<PSParameterSetInfoProxy> ParameterSetCollection { get; set; }
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

    public class PSParameterInfoProxy
    {
        public string Name { get; set; }
        public string NameLower { get; set; }
        public string ParameterType { get; set; }
        public bool IsMandatory { get; set; }
        public bool IsDynamic { get; set; }
        public int Position { get; set; }
        public bool ValueFromPipeline { get; set; }
        public string[] ValidateSetValues { get; set; }
        //public bool ValueFromPipelineByPropertyName { get; set; }
        //public bool ValueFromRemainingArguments { get; set; }
        //public string HelpMessage { get; set; }
        //public List<string> Aliases { get; set; }        
        public bool SwitchParameter { get; set; }
    }

    public class PSParameterSetInfoProxy
    {
        public PSParameterSetInfoProxy()
        {
            Parameters = new Collection<PSParameterInfoProxy>();
        }
        public string Name { get; set; }
        public string NameLower { get; set; }
        public bool IsDefault { get; set; }
        public string ToStringValue { get; set; }
        public Collection<PSParameterInfoProxy> Parameters { get; set; }

        public override string ToString()
        {
            return ToStringValue;
        }
    }

    public class PSModuleInfoProxy
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
