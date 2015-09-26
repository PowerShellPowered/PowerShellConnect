using System;
using System.Collections.Generic;


namespace PowerShellPowered.PowerShellConnect.Entities
{

    public class PSCommandHelp
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string FullHelpText { get; set; }
        public string Syntax { get; set; }
        public string Parameters { get; set; }
        public string Examples { get; set; }
        public string Category { get; set; }
        public string Component { get; set; }
        public string InputTypes { get; set; }
        public string Functionality { get; set; }
        public string NonTerminatingErrors { get; set; }
        public List<string> RelatedLinks { get; set; } = new List<string>();
        public string ReturnValues { get; set; }
        public string Role { get; set; }
        public string Synopsis { get; set; }
        public string TerminatingErrors { get; set; }
        public string Details { get; set; }

        public List<PSCommandHelpExample> PSCommandHelpExamples { get; set; } = new List<PSCommandHelpExample>();
        public List<PSCommandHelpParameter> PSCommandHelpParameters { get; set; } = new List<PSCommandHelpParameter>();
        public List<PSCommandHelpSyntax> PSCommandHelpSyntaxes { get; set; } = new List<PSCommandHelpSyntax>();

        public override string ToString()
        {
            return string.Format("Name: {0}, Detail: {1}\r\nDescription: {2}", Name, Details, Description);
        }
    }


    public class PSCommandHelpExample
    {

        public string Title { get; set; }
        public string Introduction { get; set; }
        public string Code { get; set; }
        public string Remarks { get; set; }
        public string CommandLines { get; set; }
        public string ExampleFullText { get; set; }

        public override string ToString()
        {
            return string.Format("Title: {0}, Code: {1}", Title, Code);
        }
    }


    public class PSCommandHelpParameter
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string DefaultValue { get; set; }
        public Nullable<bool> PipelineInput { get; set; }
        public Nullable<bool> Required { get; set; }
        public Nullable<bool> VariableLength { get; set; }
        public Nullable<bool> Globbing { get; set; }
        public string Position { get; set; }
        public string Type { get; set; }
        public string ParameterValue { get; set; }
        public string[] ValidateSetValues { get; set; }
        public override string ToString()
        {
            return string.Format("Name: {0}, Type: {1}", Name, ParameterValue);
        }
    }


    public class PSCommandHelpSyntax
    {
        public string Name { get; set; }
        public string Syntax { get; set; }
        public List<PSCommandHelpSyntaxParameter> PSCommandHelpSyntaxParameters { get; set; } = new List<PSCommandHelpSyntaxParameter>();

        public override string ToString()
        {
            return string.Format("Name: {0}", Name);
        }
    }


    public class PSCommandHelpSyntaxParameter
    {
        public string Name { get; set; }
        public string ParameterValue { get; set; }
        public Nullable<bool> Required { get; set; }
        public string Position { get; set; }
        public Nullable<bool> PipelineInput { get; set; }
        public string PipelineInputType { get; set; }

        public override string ToString()
        {
            return string.Format("Name: {0}, Type: {1}", Name, ParameterValue);
        }
    }
}
