using PowerShellPowered.Entities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;

namespace PowerShellPowered.PowerShellConnect
{
    public static class ExtensionMethods
    {
        public static string GetHeaderLabel(this PSObject pso)
        {
            string _r = string.Empty;
            PSPropertyInfo piDisplayName = pso.Properties["DisplayName"];
            if (piDisplayName != null && piDisplayName.Value != null)
                return piDisplayName.Value.ToString();
            PSPropertyInfo piName = pso.Properties["Name"];
            if (piName != null && piName.Value != null)
                return piName.Value.ToString();
            PSPropertyInfo piId = pso.Properties["Identity"];
            if (piId != null && piId.Value != null)
                return piId.Value.ToString();

            return _r;
        }
        public static void SetCredentialVariable(this Runspace runspace, string userName, SecureString password)
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
            SetRunspaceVariable(runspace, Constants.VariableNameStrings.Credential, cred);
        }

        public static void SetRunspaceVariable(this Runspace runspace, string name, object value)
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
            command.AddCommand(Constants.Cmdlets.SetVariable);
            command.AddParameter(Constants.ParameterNameStrings.Name, name);
            command.AddParameter(Constants.ParameterNameStrings.Value, value);
            runspace.ExecuteCommand<PSObject>(command);
        }

        public static string GetPropertyValue(this PSObject pso, string prop)
        {
            PSPropertyInfo pi = pso.Properties[prop];
            if (pi != null && pi.Value != null)
            {
                return pi.Value.ToString();
            }
            else
                return string.Empty;
        }

        public static string GetPropertyValueNested(this PSObject pso, string prop, string propNested)
        {
            throw new NotImplementedException();
            //PSPropertyInfo pi = pso.Properties[prop];
            //if (pi != null && pi.Value != null)
            //{
            //    return pi.Value.ToString();
            //}
            //else
            //    return string.Empty;
        }

        public static bool TryGetPropertyAsPSObject(this PSObject psObject, string prop, out PSObject psOutObject)
        {
            return TryGetPropertyAsTObject<PSObject>(psObject, prop, out psOutObject);
        }

        public static bool TryGetPropertyAsTObject<T>(this PSObject psObject, string prop, out T outObj)
        {
            PSPropertyInfo pi = psObject.Properties[prop];
            if (pi != null && pi.Value != null && pi.Value is T)
            {
                outObj = (T)pi.Value;
                return true;
            }
            else
            {
                outObj = default(T);
                return false;
            }
        }

        public static bool IsArrayList(this PSObject pso)
        {
            return (pso.BaseObject is IList);
        }

        public static ArrayList ToArrayList(this PSObject pso)
        {
            return pso.BaseObject as ArrayList;
        }

        public static IList ToIList(this PSObject pso)
        {
            return pso.BaseObject as IList;
        }

        public static bool IsHashTable(this PSObject pso)
        {
            return (pso.BaseObject is IDictionary);
        }

        public static Hashtable ToHashTable(this PSObject pso)
        {
            return pso.BaseObject as Hashtable;
        }

        public static bool IsPSObject(this PSObject pso)
        {
            return (pso.BaseObject is PSObject);
        }

        public static PSObject ToPSObject(this PSObject pso)
        {
            return pso.BaseObject as PSObject;
        }

        public static Collection<PSObject> ExecuteCommand(this Runspace runspace, PSCommand command, Action<PSDataStreams> psDataStreamAction = null)
        {
            return ExecuteCommand<PSObject>(runspace, command, psDataStreamAction);
        }

        public static Collection<T> ExecuteCommand<T>(this Runspace runspace, PSCommand command, Action<PSDataStreams> psDataStreamAction = null)
        {
            if (command == null)
            {
                throw new ArgumentOutOfRangeException("command");
            }

            if (runspace == null)
            {
                throw new ArgumentOutOfRangeException("runspace");
            }

            using (var ps = PowerShell.Create())
            {
                ps.Commands = command;
                ps.Runspace = runspace;

                if (psDataStreamAction != null)
                    psDataStreamAction(ps.Streams);

                var result = ps.Invoke<T>();

                return result;
            }
        }

        public static Collection<PSObject> ExecuteScript(this Runspace runspace, string scriptText, Dictionary<string, object> scriptParameters = null, bool out_String = false, bool stream = false, int width = 10000, Action<PSDataStreams> psDataStreamAction = null)
        {
            return ExecuteScript<PSObject>(runspace, scriptText, scriptParameters, out_String, stream, width, psDataStreamAction);
        }

        public static Collection<T> ExecuteScript<T>(this Runspace runspace, string scriptText, Dictionary<string, object> scriptParameters = null, bool out_String = false, bool stream = false, int width = 10000, Action<PSDataStreams> psDataStreamAction = null)
        {

            if (scriptText == null)
            {
                throw new ArgumentOutOfRangeException("scriptText");
            }

            using (PowerShell ps = PowerShell.Create())
            {
                ps.Runspace = runspace;
                ps.AddScript(scriptText);

                if (scriptParameters != null && scriptParameters.Count > 0) ps.AddParameters(scriptParameters);

                if (out_String && typeof(T).Equals(typeof(string)))
                {
                    ps.AddCommand("Out-String").AddParameter("Width", width);

                    if (stream) ps.AddParameter("Stream");

                }
                //ps.AddCommand("Out-String");
                if (psDataStreamAction != null)
                    psDataStreamAction(ps.Streams);
                var result = ps.Invoke<T>();
                return result;
            }
        }
        
        public static CmdInfo ToCmdInfo(this CommandInfo commandInfo, bool includeCommonParameters = false)
        {
            CmdInfo _res = new CmdInfo();

            CommandTypes commandType = commandInfo.CommandType;

            _res.IsRemoteCommand = false;
            _res.Name = commandInfo.Name;
            _res.NameLower = commandInfo.Name.ToLower();
            _res.CommandType = commandType.ToString();
            _res.Definition = commandInfo.Definition.Trim();
            _res.Module = commandInfo.Module == null ? string.Empty : commandInfo.Module.Name;
            _res.ModuleName = commandInfo.ModuleName;
            if (commandInfo.OutputType != null)
                _res.OutputType = (from a in commandInfo.OutputType select a.Name).ToList();
            if (commandInfo.ParameterSets != null)
            {
                _res.ParameterSets = (from a in commandInfo.ParameterSets select a.ToString().Trim()).ToList();
                _res.CmdParameterSets = new List<CmdParameterSetInfo>();
                commandInfo.ParameterSets.ToList().AsParallel().ForAll(x =>
                    {
                        CmdParameterSetInfo psip = new CmdParameterSetInfo();
                        psip.IsDefault = x.IsDefault;
                        psip.Name = x.Name;
                        psip.NameLower = x.Name.ToLower();
                        psip.ToStringValue = x.ToString();
                        x.Parameters.ToList().AsParallel().ForAll(y =>
                            {
                                if (!Constants.CommomParamaters.Contains(y.Name) || includeCommonParameters)
                                {
                                    CmdParameterInfo p = new CmdParameterInfo();
                                    p.Name = y.Name;
                                    p.ParameterType = y.ParameterType.ToString();
                                    p.IsMandatory = y.IsMandatory;
                                    p.IsDynamic = y.IsDynamic;
                                    p.Position = y.Position;
                                    p.ValueFromPipeline = y.ValueFromPipeline;
                                    p.SwitchParameter = y.ParameterType.ToString().Contains("SwitchParameter");

                                    if (y.ParameterType.IsEnum)
                                        p.ValidateSetValues = Enum.GetNames(y.ParameterType).ToList();

                                    lock (psip.Parameters) { psip.Parameters.Add(p); }
                                }
                            });

                        lock (_res.CmdParameterSets) { _res.CmdParameterSets.Add(psip); }
                    });
            }

            if (commandInfo.Parameters != null)
            {
                if (includeCommonParameters)
                    _res.Parameters = (from a in commandInfo.Parameters select a.Key).ToList();
                else
                    _res.Parameters = (from a in commandInfo.Parameters where !Constants.CommomParamaters.Contains(a.Key) select a.Key).ToList();

                _res.CmdParameters = new List<CmdParameterInfo>();

                //Parallel.ForEach(commandInfo.Parameters.ToList(), (x) => { },,);


                commandInfo.Parameters.ToList().AsParallel().ForAll(x =>
                {
                    if (!Constants.CommomParamaters.Contains(x.Key) || includeCommonParameters)
                    {
                        CmdParameterInfo pip = new CmdParameterInfo();
                        pip.Name = x.Key;
                        pip.ParameterType = x.Value.ParameterType.Name;
                        pip.IsDynamic = x.Value.IsDynamic;
                        pip.SwitchParameter = x.Value.SwitchParameter;
                        var atr = x.Value.Attributes.FirstOrDefault((a) => a.TypeId.ToString().Equals("System.Management.Automation.ParameterAttribute")) as ParameterAttribute;
                        if (atr != null)
                        {
                            pip.IsMandatory = atr.Mandatory;
                            pip.Position = atr.Position;
                            pip.ValueFromPipeline = atr.ValueFromPipeline;
                        }

                        if (x.Value.ParameterType.IsEnum)
                            pip.ValidateSetValues = Enum.GetNames(x.Value.ParameterType).ToList();

                        lock (_res.CmdParameters) { _res.CmdParameters.Add(pip); }

                    }
                });
            }


            _res.RemotingCapability = RemotingCapability.None.ToString();
            try
            {
                if (commandInfo.RemotingCapability != null)
                    _res.RemotingCapability = commandInfo.RemotingCapability.ToString();
            }
            catch (Exception) { }




            //if (commandType != CommandTypes.ExternalScript && commandInfo.Visibility != null)
            _res.Visibility = commandInfo.Visibility.ToString();

            dynamic commandInfoDynamic = commandInfo;
            switch (commandType)
            {
                case CommandTypes.Alias:
                    _res.Noun = commandInfoDynamic.Noun;
                    if (commandInfoDynamic.Options != null) _res.Options = commandInfoDynamic.Options.ToString();
                    _res.Description = commandInfoDynamic.Description;
                    if (commandInfoDynamic.ReferencedCommand != null) _res.ReferencedCommand = commandInfoDynamic.ReferencedCommand.ToString();
                    if (commandInfoDynamic.ResolvedCommand != null) _res.ResolvedCommand = commandInfoDynamic.ResolvedCommand.ToString();
                    break;
                case CommandTypes.Cmdlet:
                    _res.Noun = commandInfoDynamic.Noun;
                    if (commandInfoDynamic.Options != null) _res.Options = commandInfoDynamic.Options.ToString();
                    _res.DefaultParameterSet = commandInfoDynamic.DefaultParameterSet;
                    _res.HelpFile = commandInfoDynamic.HelpFile;
                    //if (commandInfoDynamic.HelpUri != null) _res.HelpUri = commandInfoDynamic.HelpUri.ToString();
                    _res.Verb = commandInfoDynamic.Verb;
                    if (commandInfoDynamic.ImplementingType != null) _res.ImplementingType = commandInfoDynamic.ImplementingType.FullName;
                    if (commandInfoDynamic.PSSnapIn != null) _res.PSSnapIn = commandInfoDynamic.PSSnapIn.Name;
                    break;
                case CommandTypes.ExternalScript:
                    if (commandInfoDynamic.OriginalEncoding != null) _res.OriginalEncoding = commandInfoDynamic.OriginalEncoding.ToString();
                    _res.Path = commandInfoDynamic.Path;
                    if (commandInfoDynamic.ScriptBlock != null) _res.ScriptBlock = commandInfoDynamic.ScriptBlock.ToString();
                    _res.ScriptContents = commandInfoDynamic.ScriptContents;
                    break;
                case CommandTypes.Function:
                    _res.Noun = commandInfoDynamic.Noun;
                    if (commandInfoDynamic.Options != null) _res.Options = commandInfoDynamic.Options.ToString();
                    _res.DefaultParameterSet = commandInfoDynamic.DefaultParameterSet;
                    _res.HelpFile = commandInfoDynamic.HelpFile;
                    //if (commandInfoDynamic.HelpUri != null) _res.HelpUri = commandInfoDynamic.HelpUri.ToString();
                    _res.Verb = commandInfoDynamic.Verb;

                    _res.CmdletBinding = commandInfoDynamic.CmdletBinding;
                    _res.Description = commandInfoDynamic.Description;
                    if (commandInfoDynamic.ScriptBlock != null) _res.ScriptBlock = commandInfoDynamic.ScriptBlock.ToString();
                    break;
                case CommandTypes.Script:
                    if (commandInfoDynamic.ScriptBlock != null) _res.ScriptBlock = commandInfoDynamic.ScriptBlock.ToString();
                    break;
                default:
                    break;
            }

            return _res;
        }
        
        public static ModuleInfo ToModuleInfo(this PSModuleInfo moduleInfo)
        {
            ModuleInfo _res = new ModuleInfo()
            {
                Guid = moduleInfo.Guid,
                Name = moduleInfo.Name,
                Description = moduleInfo.Description,
                NestedModules = moduleInfo.NestedModules.Select<PSModuleInfo, string>(p => p.Name).ToList(),
                RequiredModules = moduleInfo.RequiredModules.Select<PSModuleInfo, string>(p => p.Name).ToList(),
                ModuleType = moduleInfo.ModuleType.ToString(),
                RootModule = moduleInfo.RootModule,
                CompanyName = moduleInfo.CompanyName,
                Author = moduleInfo.Author,
                HelpInfoUri = moduleInfo.HelpInfoUri
            };

            return _res;
        }
        
    }
}
