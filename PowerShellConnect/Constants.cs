

using System.Collections.Generic;
namespace PowerShellPowered.PowerShellConnect
{
    public static class Constants
    {
        public static List<string> CommomParamaters = new List<string>() { "Debug", "ErrorAction", "ErrorVariable", "OutVariable", "OutBuffer", "Verbose", "WarningAction", "WarningVariable", "WhatIf" };

        public static class PSCmdlets
        {
            internal const string SetVariable = "Set-Variable";
            internal const string GetCommand = "Get-Command";
            public static string GetHelp = "Get-Help";

        }

        public static class PSSessionScripts
        {
            internal const string RemovePSSession = "Remove-PSSession";
            internal const string ImportPSSession = "Import-PSSession";

            internal const string AllowRedirectionInNewPSSession = " -AllowRedirection";

            internal const string NewPSSessionScriptWithBasicAuth = "New-PSSession -ConfigurationName:{0} -Authentication:Basic -ConnectionUri:{1} -Credential $" + Constants.PSVariableNameStrings.Credential;

            internal const string NewPSSessionScriptWithDefaultAuth = "New-PSSession -ConfigurationName:{0} -ConnectionUri:{1} -Credential $" + Constants.PSVariableNameStrings.Credential;

        }

        internal static class PSHelpParsingScripts
        {
            internal const string GetCommandScriptOld = @"

param(
[System.Management.Automation.Runspaces.PSSession]
$Session,
[string]
$Name = '*',
[int]
$CommandType = 91,
[int]
$TotalCount = -1,
[string]
$Module
)

# CommandTypes
#Alias
#1
#Function
#2
#Filter
#4
#Cmdlet
#8
#ExternalScript
#16
#Application
#32
#Script
#64
#Workflow
#128
#All
#255


#Write-Verbose $Session -Verbose
#Write-Verbose $Name -Verbose

$ScriptBlockString = 'Get-Command -Name $args[0] -CommandType $args[1] -TotalCount $args[2] '


if($Module -ne $null -or $Module -ne '')
{ $ScriptBlockString += ' -Module $args[3]'}

$ScriptBlock = [System.Management.Automation.ScriptBlock]::Create($ScriptBlockString)

Write-Verbose $ScriptBlockString -Verbose
Write-Verbose $ScriptBlock.ToString() -Verbose


$CommandInfoProxyArray = @()

if($Session -ne $null)
{
    $CommandInfos = Invoke-Command -Session $Session -ScriptBlock $ScriptBlock -ArgumentList $Name,$CommandType,$TotalCount,$Module
}
else
{
    $CommandInfos = Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Name,$CommandType,$TotalCount,$Module
}
    
Write-Verbose $CommandInfos.Count -Verbose

foreach($CommandInfo in $CommandInfos)
{
    Write-Verbose $CommandInfo.Name -Verbose
    Write-Verbose [string]$CommandInfo.CommandType -Verbose
    #Write-Verbose $CommandInfo.OriginalEncoding -Verbose
    #Write-Verbose $CommandInfo.ToString() -Verbose
    $commandType = [System.Management.Automation.CommandTypes]$CommandInfo.CommandType

    $CommandInfoProxy = New-Object -TypeName PSObject -Property @{'IsRemoteCommand' = [bool]$true; 'CmdletBinding' = [bool]$false;[string]'CommandType'=$null;'DefaultParameterSet'=$null;'Definition'=$null;'Description'=$null;'HelpFile'=$null;'HelpUri'=$null;'ImplementingType'=$null;'Module'=$null;'ModuleName'=$null;'Name'=$null;'Noun'=$null;'Options'=$null;'OriginalEncoding'=$null;'OutputType'=$null;'Parameters'=$null;'ParameterSets'=$null;'Path'=$null;'PSSnapIn'=$null;'ReferencedCommand'=$null;'RemotingCapability'=$null;'ResolvedCommand'=$null;'ScriptBlock'=$null;'ScriptContents'=$null;'Verb'=$null;'Visibility'=$null;}
        
    $CommandInfoProxy.IsRemoteCommand = -not($CommandInfo -is [System.Management.Automation.CommandInfo])
    $CommandInfoProxy.Name = $CommandInfo.Name
    $CommandInfoProxy.CommandType = [string]$CommandInfo.CommandType
    $CommandInfoProxy.Definition = $CommandInfo.Definition
    if($CommandInfo.Module){$CommandInfoProxy.Module = $CommandInfo.Module.Name}
    $CommandInfoProxy.ModuleName = $CommandInfo.ModuleName
    [string[]]$parameterssets = 
    if($CommandInfo.ParameterSets){[string[]]$CommandInfoProxy.ParameterSets  = $CommandInfo.ParameterSets | %{$_.ToString()}}
    [string[]]$parameters = @()
    ($CommandInfo.Parameters).Keys | %{$parameters += $_.ToString()}
    if($CommandInfo.Parameters){$CommandInfoProxy.Parameters  = $parameters}
    #$CommandInfoProxy.RemotingCapability  = [System.Management.Automation.RemotingCapability]$CommandInfo.RemotingCapability
    if($CommandInfo.RemotingCapability){$CommandInfoProxy.RemotingCapability  = $CommandInfo.RemotingCapability.ToString()}

    #$CommandInfoProxy.OutputType = $CommandInfo.OutputType
    if($CommandInfo.OutputType){ $CommandInfoProxy.OutputType  = ($CommandInfo.OutputType  | %{$_.ToString()}) }
    
    #$CommandInfoProxy.Visibility = [System.Management.Automation.SessionStateEntryVisibility]$CommandInfo.Visibility
    if($CommandInfo.Visibility){$CommandInfoProxy.Visibility = $CommandInfo.Visibility.ToString()}
    
        
    if( $commandType -eq [System.Management.Automation.CommandTypes]::Alias)
    {
        if($CommandInfo.Noun){$CommandInfoProxy.Noun  = $CommandInfo.Noun}
        #$CommandInfoProxy.Options = [System.Management.Automation.ScopedItemOptions]$CommandInfo.Options
        if($CommandInfo.Options){$CommandInfoProxy.Options = $CommandInfo.Options.ToString()}
        $CommandInfoProxy.Description  = $CommandInfo.Description            

        if($CommandInfo.ReferencedCommand){$CommandInfoProxy.ReferencedCommand  = $CommandInfo.ReferencedCommand.ToString()}
        if($CommandInfo.ResolvedCommand){$CommandInfoProxy.ResolvedCommand  = $CommandInfo.ResolvedCommand.ToString()}
    }

        
    if( $commandType -eq [System.Management.Automation.CommandTypes]::Cmdlet)
    {
        if($CommandInfo.Noun){$CommandInfoProxy.Noun  = $CommandInfo.Noun}
        #$CommandInfoProxy.Options = [System.Management.Automation.ScopedItemOptions]$CommandInfo.Options
        if($CommandInfo.Options){$CommandInfoProxy.Options = $CommandInfo.Options.ToString()}

        $CommandInfoProxy.DefaultParameterSet  = $CommandInfo.DefaultParameterSet            
        $CommandInfoProxy.HelpFile  = $CommandInfo.HelpFile
        if($CommandInfo.HelpUri){$CommandInfoProxy.HelpUri  = $CommandInfo.HelpUri.ToString().Trim()}
        $CommandInfoProxy.Verb  = $CommandInfo.Verb        

        if($CommandInfo.ImplementingType){$CommandInfoProxy.ImplementingType  = $CommandInfo.ImplementingType.ToString()}
        if($CommandInfo.PSSnapIn){$CommandInfoProxy.PSSnapIn  = $CommandInfo.PSSnapIn.ToString()}
    }

    if( $commandType -eq [System.Management.Automation.CommandTypes]::ExternalScript)
    {        
        if($CommandInfo.ScriptBlock){$CommandInfoProxy.ScriptBlock  = $CommandInfo.ScriptBlock.ToString()}
        
        if($CommandInfo.OriginalEncoding){$CommandInfoProxy.OriginalEncoding  = $CommandInfo.OriginalEncoding.ToString()}
        $CommandInfoProxy.Path  = $CommandInfo.Path
        $CommandInfoProxy.ScriptContents  = $CommandInfo.ScriptContents
    }
    
    if( $commandType -eq [System.Management.Automation.CommandTypes]::Function)
    {
        if($CommandInfo.Noun){$CommandInfoProxy.Noun  = $CommandInfo.Noun}
        #$CommandInfoProxy.Options = [System.Management.Automation.ScopedItemOptions]$CommandInfo.Options
        if($CommandInfo.Options){$CommandInfoProxy.Options = $CommandInfo.Options.ToString()}

        if($CommandInfo.ScriptBlock){$CommandInfoProxy.ScriptBlock  = $CommandInfo.ScriptBlock.ToString()}

        $CommandInfoProxy.DefaultParameterSet  = $CommandInfo.DefaultParameterSet            
        $CommandInfoProxy.HelpFile  = $CommandInfo.HelpFile
        if($CommandInfo.HelpUri){$CommandInfoProxy.HelpUri  = $CommandInfo.HelpUri.ToString()}
        $CommandInfoProxy.Verb  = $CommandInfo.Verb        

        $CommandInfoProxy.Description  = $CommandInfo.Description            

        $CommandInfoProxy.CmdletBinding  = [bool]$CommandInfo.CmdletBinding            
    }

    
    if( $commandType -eq [System.Management.Automation.CommandTypes]::Script)
    {
        
        if($CommandInfo.ScriptBlock){$CommandInfoProxy.ScriptBlock  = $CommandInfo.ScriptBlock.ToString()}
    }

        
    $CommandInfoProxyArray += $CommandInfoProxy
}
return $CommandInfoProxyArray
";

            internal const string GetCommandScript = @"
param(
[System.Management.Automation.Runspaces.PSSession]
$Session,
[string]
$Name = '*',
[int]
$TotalCount = -1,
[int]
$CommandType = 91,
[string]
$EntitiesFile = '',
[string]
$Module
)

# CommandTypes
#Alias
#1
#Function
#2
#Filter
#4
#Cmdlet
#8
#ExternalScript
#16
#Application
#32
#Script
#64
#Workflow
#128
#All
#255


#Write-Verbose $Session -Verbose
#Write-Verbose $Name -Verbose

$ScriptBlockString = 'Get-Command -Name $args[0] -CommandType $args[1] -TotalCount $args[2] '


if($Module -ne $null -or $Module -ne '')
{ $ScriptBlockString += ' -Module $args[3]'}

$ScriptBlock = [System.Management.Automation.ScriptBlock]::Create($ScriptBlockString)

Write-Verbose $TotalCount -Verbose
#Write-Verbose $ScriptBlock.ToString() -Verbose


$CommandInfoProxyArray = @()

if($Session -ne $null)
{
    $CommandInfos = Invoke-Command -Session $Session -ScriptBlock $ScriptBlock -ArgumentList $Name,$CommandType,$TotalCount,$Module
}
else
{
    $CommandInfos = Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Name,$CommandType,$TotalCount,$Module
}
    
#Write-Verbose $CommandInfos.Count -Verbose

$useType = $null

if($EntitiesFile -ne $null -and  (Test-Path $EntitiesFile))
{
    $useType = [System.Reflection.Assembly]::LoadFile($EntitiesFile)
}


foreach($CommandInfo in $CommandInfos)
{
    #Write-Verbose $CommandInfo.Name -Verbose
    #Write-Verbose [string]$CommandInfo.CommandType -Verbose
    #Write-Verbose $CommandInfo.OriginalEncoding -Verbose
    #Write-Verbose $CommandInfo.ToString() -Verbose
    $commandType = [System.Management.Automation.CommandTypes]$CommandInfo.CommandType


    if($useType)
    {
        #Write-Verbose 'using Types' -Verbose
        $CommandInfoProxy = New-Object -TypeName PowerShellPowered.PowerShellConnect.Entities.PSCommandInfoProxy
    }
    else
    {
        $CommandInfoProxy = New-Object -TypeName PSObject -Property @{'IsRemoteCommand' = [bool]$true; 'CmdletBinding' = [bool]$false;[string]'CommandType'=$null;'DefaultParameterSet'=$null;'Definition'=$null;'Description'=$null;'HelpFile'=$null;'HelpUri'=$null;'ImplementingType'=$null;'Module'=$null;'ModuleName'=$null;'Name'=$null;'NameLower'=$null;'Noun'=$null;'Options'=$null;'OriginalEncoding'=$null;'OutputType'=$null;'Parameters'=$null;'ParameterSets'=$null;'Path'=$null;'PSSnapIn'=$null;'ReferencedCommand'=$null;'RemotingCapability'=$null;'ResolvedCommand'=$null;'ScriptBlock'=$null;'ScriptContents'=$null;'Verb'=$null;'Visibility'=$null;}
    }
    
        
    $CommandInfoProxy.IsRemoteCommand = -not($CommandInfo -is [System.Management.Automation.CommandInfo])
    $CommandInfoProxy.Name = $CommandInfo.Name
    $CommandInfoProxy.NameLower = $CommandInfo.Name.ToLower()
    $CommandInfoProxy.CommandType = [string]$CommandInfo.CommandType
    $CommandInfoProxy.Definition = $CommandInfo.Definition
    if($CommandInfo.Module){$CommandInfoProxy.Module = $CommandInfo.Module.Name}
    $CommandInfoProxy.ModuleName = $CommandInfo.ModuleName
    [string[]]$parameterssets = @()
    if($CommandInfo.ParameterSets){[string[]]$CommandInfoProxy.ParameterSets  = $CommandInfo.ParameterSets | %{$_.ToString()}}    
    [string[]]$parameters = @()
    if($CommandInfo.Parameters){($CommandInfo.Parameters).Keys | %{$parameters += $_.ToString()};$CommandInfoProxy.Parameters  = $parameters}
    #$CommandInfoProxy.RemotingCapability  = [System.Management.Automation.RemotingCapability]$CommandInfo.RemotingCapability
    if($CommandInfo.RemotingCapability){$CommandInfoProxy.RemotingCapability  = $CommandInfo.RemotingCapability.ToString()}

    #$CommandInfoProxy.OutputType = $CommandInfo.OutputType
    if($CommandInfo.OutputType){ $CommandInfoProxy.OutputType  = ($CommandInfo.OutputType  | %{$_.ToString()}) }
    
    #$CommandInfoProxy.Visibility = [System.Management.Automation.SessionStateEntryVisibility]$CommandInfo.Visibility
    if($CommandInfo.Visibility){$CommandInfoProxy.Visibility = $CommandInfo.Visibility.ToString()}
    
        
    if( $commandType -eq [System.Management.Automation.CommandTypes]::Alias)
    {
        if($CommandInfo.Noun){$CommandInfoProxy.Noun  = $CommandInfo.Noun}
        #$CommandInfoProxy.Options = [System.Management.Automation.ScopedItemOptions]$CommandInfo.Options
        if($CommandInfo.Options){$CommandInfoProxy.Options = $CommandInfo.Options.ToString()}
        $CommandInfoProxy.Description  = $CommandInfo.Description            

        if($CommandInfo.ReferencedCommand){$CommandInfoProxy.ReferencedCommand  = $CommandInfo.ReferencedCommand.ToString()}
        if($CommandInfo.ResolvedCommand){$CommandInfoProxy.ResolvedCommand  = $CommandInfo.ResolvedCommand.ToString()}
    }

        
    if( $commandType -eq [System.Management.Automation.CommandTypes]::Cmdlet)
    {
        if($CommandInfo.Noun){$CommandInfoProxy.Noun  = $CommandInfo.Noun}
        #$CommandInfoProxy.Options = [System.Management.Automation.ScopedItemOptions]$CommandInfo.Options
        if($CommandInfo.Options){$CommandInfoProxy.Options = $CommandInfo.Options.ToString()}

        $CommandInfoProxy.DefaultParameterSet  = $CommandInfo.DefaultParameterSet            
        $CommandInfoProxy.HelpFile  = $CommandInfo.HelpFile
        if($CommandInfo.HelpUri){$CommandInfoProxy.HelpUri  = $CommandInfo.HelpUri.ToString().Trim()}
        $CommandInfoProxy.Verb  = $CommandInfo.Verb        

        if($CommandInfo.ImplementingType){$CommandInfoProxy.ImplementingType  = $CommandInfo.ImplementingType.ToString()}
        if($CommandInfo.PSSnapIn){$CommandInfoProxy.PSSnapIn  = $CommandInfo.PSSnapIn.ToString()}
    }

    if( $commandType -eq [System.Management.Automation.CommandTypes]::ExternalScript)
    {        
        if($CommandInfo.ScriptBlock){$CommandInfoProxy.ScriptBlock  = $CommandInfo.ScriptBlock.ToString()}
        
        if($CommandInfo.OriginalEncoding){$CommandInfoProxy.OriginalEncoding  = $CommandInfo.OriginalEncoding.ToString()}
        $CommandInfoProxy.Path  = $CommandInfo.Path
        $CommandInfoProxy.ScriptContents  = $CommandInfo.ScriptContents
    }
    
    if( $commandType -eq [System.Management.Automation.CommandTypes]::Function)
    {
        if($CommandInfo.Noun){$CommandInfoProxy.Noun  = $CommandInfo.Noun}
        #$CommandInfoProxy.Options = [System.Management.Automation.ScopedItemOptions]$CommandInfo.Options
        if($CommandInfo.Options){$CommandInfoProxy.Options = $CommandInfo.Options.ToString()}

        if($CommandInfo.ScriptBlock){$CommandInfoProxy.ScriptBlock  = $CommandInfo.ScriptBlock.ToString()}

        $CommandInfoProxy.DefaultParameterSet  = $CommandInfo.DefaultParameterSet            
        $CommandInfoProxy.HelpFile  = $CommandInfo.HelpFile
        if($CommandInfo.HelpUri){$CommandInfoProxy.HelpUri  = $CommandInfo.HelpUri.ToString()}
        $CommandInfoProxy.Verb  = $CommandInfo.Verb        

        $CommandInfoProxy.Description  = $CommandInfo.Description            

        $CommandInfoProxy.CmdletBinding  = [bool]$CommandInfo.CmdletBinding            
    }

    
    if( $commandType -eq [System.Management.Automation.CommandTypes]::Script)
    {
        
        if($CommandInfo.ScriptBlock){$CommandInfoProxy.ScriptBlock  = $CommandInfo.ScriptBlock.ToString()}
    }

    #$CommandInfoProxy
    $CommandInfoProxyArray += $CommandInfoProxy
}
return $CommandInfoProxyArray
";

            internal const string GetHelpScriptOld = @"

param(
[System.Management.Automation.Runspaces.PSSession]
$Session,
[string]
$CommandName
)

function getparameterdetail
{
param($type, $parameter)

if($type.name -eq 'type'){$output = ($parameter.$($type.Name)).name}else{ $output = $parameter.$($type.name)}
#Write-Verbose ($p | Out-String) -Verbose
($output | Out-String  -Width 10000).Trim()
}

if($Session -ne $null)
{
    $help = Invoke-Command -Session $Session -ScriptBlock {get-help $args[0] -Full} -ArgumentList $CommandName
}
else
{
    $help = get-help $CommandName -Full
}

if($help -eq $null){return $null}
$h = New-Object -TypeName psobject
$h | Add-Member -MemberType NoteProperty -Name alertSet -Value ($help.alertSet | Out-String -Width 10000).Trim()
$h | Add-Member -MemberType NoteProperty -Name Category -Value ($help.Category | Out-String -Width 10000).Trim()
$h | Add-Member -MemberType NoteProperty -Name Component -Value ($help.Component | Out-String -Width 10000).Trim()
$h | Add-Member -MemberType NoteProperty -Name description -Value ($help.description | Out-String -Width 10000).Trim()
$h | Add-Member -MemberType NoteProperty -Name details -Value ($help.details | Out-String -Width 10000).Trim()
$h | Add-Member -MemberType NoteProperty -Name examples -Value ($help.examples | Out-String -Width 10000).Trim()
$h | Add-Member -MemberType NoteProperty -Name Functionality -Value ($help.Functionality | Out-String -Width 10000).Trim()
$h | Add-Member -MemberType NoteProperty -Name inputTypes -Value ($help.inputTypes | Out-String -Width 10000).Trim()
$h | Add-Member -MemberType NoteProperty -Name ModuleName -Value ($help.ModuleName | Out-String -Width 10000).Trim()
$h | Add-Member -MemberType NoteProperty -Name Name -Value ($help.Name | Out-String -Width 10000).Trim()
$h | Add-Member -MemberType NoteProperty -Name nonTerminatingErrors -Value ($help.nonTerminatingErrors | Out-String -Width 10000).Trim()
$h | Add-Member -MemberType NoteProperty -Name parameters -Value ($help.parameters | Out-String -Width 10000).Trim()
$h | Add-Member -MemberType NoteProperty -Name PSSnapIn -Value ($help.PSSnapIn | Out-String -Width 10000).Trim()
$h | Add-Member -MemberType NoteProperty -Name relatedLinks -Value ($help.relatedLinks | Out-String -Width 10000).Trim()
$h | Add-Member -MemberType NoteProperty -Name returnValues -Value ($help.returnValues | Out-String -Width 10000).Trim()
$h | Add-Member -MemberType NoteProperty -Name Role -Value ($help.Role | Out-String -Width 10000).Trim()
$h | Add-Member -MemberType NoteProperty -Name Synopsis -Value ($help.Synopsis | Out-String -Width 10000).Trim()
$h | Add-Member -MemberType NoteProperty -Name syntax -Value ($help.syntax | Out-String -Width 10000).Trim()
$h | Add-Member -MemberType NoteProperty -Name terminatingErrors -Value ($help.terminatingErrors | Out-String -Width 10000).Trim()
$h | Add-Member -MemberType NoteProperty -Name FullHelpText -Value ($help | Out-String -Width 10000).Trim()

[PSObject[]]$params = @()
foreach($p in $help.parameters.parameter)
{
    $pi = New-Object -TypeName PSObject
    $ti = Get-Member -InputObject $p -MemberType Properties
    foreach($t in $ti){        
        $pi | Add-Member -MemberType NoteProperty -Name ($t.Name | Out-String).Trim() -Value (getparameterdetail $t $p | out-string -Width 10000).trim()
    }
    $params +=$pi
}
#$px = New-Object -TypeName PSObject
#$px | Add-Member -MemberType NoteProperty -Name parameter -Value $params
#$h | Add-Member -MemberType NoteProperty -Name ParametersEx -Value $px
$h | Add-Member -MemberType NoteProperty -Name ParametersEx -Value $params

[PSObject[]]$examples = @()
foreach($e in $help.examples.example)
{
    $ex = New-Object -TypeName PSObject
    $ex | Add-Member -MemberType NoteProperty -Name code -Value ($e.code | Out-String -Width 10000).Trim()
    $ex | Add-Member -MemberType NoteProperty -Name commandLines -Value ($e.commandLines.commandLine.commandText | Out-String -Width 10000).Trim()
    $ex | Add-Member -MemberType NoteProperty -Name introduction -Value ($e.introduction | Out-String -Width 10000).Trim()
    $ex | Add-Member -MemberType NoteProperty -Name remarks -Value ($e.remarks | Out-String -Width 10000).Trim()
    $ex | Add-Member -MemberType NoteProperty -Name title -Value ($e.title | Out-String -Width 10000).Trim()
    $examples +=$ex
}
$h | Add-Member -MemberType NoteProperty -Name ExamplesEx -Value $examples

[PSObject[]]$syntaxes = @()

foreach($s in $help.syntax.syntaxItem)
{
    $sx = New-Object -TypeName PSObject
    $sx | Add-Member -MemberType NoteProperty -Name name -Value ($s.name | Out-String -Width 10000).Trim()

    [PSObject[]]$sxis = @()       

    $i=0;
    foreach($sp in $s.parameter)
    {
        $sxi = New-Object -TypeName PSObject
        $sxi | Add-Member -MemberType NoteProperty -Name aliases -Value ($sp.aliases | Out-String -Width 10000).Trim()
        $sxi | Add-Member -MemberType NoteProperty -Name description -Value ($sp.description | Out-String -Width 10000).Trim()
        $sxi | Add-Member -MemberType NoteProperty -Name globbing -Value ($sp.globbing | Out-String -Width 10000).Trim()
        $sxi | Add-Member -MemberType NoteProperty -Name name -Value ($sp.name | Out-String -Width 10000).Trim()
        $sxi | Add-Member -MemberType NoteProperty -Name parameterValue -Value ($sp.parameterValue | Out-String -Width 10000).Trim()
        $sxi | Add-Member -MemberType NoteProperty -Name pipelineInput -Value ($sp.pipelineInput | Out-String -Width 10000).Trim()
        $sxi | Add-Member -MemberType NoteProperty -Name position -Value ($sp.position | Out-String -Width 10000).Trim()
        $sxi | Add-Member -MemberType NoteProperty -Name required -Value ($sp.required | Out-String -Width 10000).Trim()
        $sxi | Add-Member -MemberType NoteProperty -Name variableLength -Value ($sp.variableLength | Out-String -Width 10000).Trim()
        $sxis+=$sxi
    }
    $sx | Add-Member -MemberType NoteProperty -Name parameters -Value $sxis        
    $syntaxes +=$sx
}
$h | Add-Member -MemberType NoteProperty -Name SyntaxEx -Value $syntaxes
$h
";
            internal const string GetHelpScript = @"
param(
[System.Management.Automation.Runspaces.PSSession]
$Session,
[string]
$Name,
[string]
$EntitiesFile = ''
)

    function getparameterdetail
    {
        param($type, $parameter)

        if($type.name -eq 'type'){$output = ($parameter.$($type.Name)).name}else{ $output = $parameter.$($type.name)}
        #Write-Verbose ($p | Out-String) -Verbose
        ($output | Out-String  -Width 10000).Trim()
    }

    function ParseBool
    {
        param([string]$str, [bool]$indexOfTrue = $true, [bool] $nullable = $true)
        $outnull = $null
        if ($nullable -and [string]::IsNullOrEmpty($str)) {return $null}
        if ([bool]::TryParse($str, [ref]$outnull))
            {return [bool]::Parse($str)}
        elseif ($str.Equals('$true', [System.StringComparison]::OrdinalIgnoreCase))
            {return $true}
        elseif ($indexOfTrue -and ($str.IndexOf('true',[System.StringComparison]::OrdinalIgnoreCase) -ge 0))
            {return $true}
        elseif ($str.Equals('0'))
            {return $false}
        elseif ($str.Equals('1'))
            {return $true}
        elseif ($nullable)
            {return $null}
        else
            {return $false}
    }

    function GetNavigationLinks
    {
        param($help)
        
        $link = @()
        if(-not $help.relatedLinks) {return $link}
        
        $linkObjs = $help.relatedLinks.navigationLink | ?{!$_.linktext -or $_.linktext -like 'online version*'}

        if($linkObjs)
        {
            $linkObjs | %{$link += $_.uri}
        }

        if($link.count -eq 0)
        {
            $link += ($help.relatedLinks | Out-String -Width 10000).Trim()
        }

        return $link
    }
    
    if($Session -ne $null)
    {
        $help = Invoke-Command -Session $Session -ScriptBlock {get-help $args[0] -Full} -ArgumentList $Name
    }
    else
    {
        $help = get-help $Name -Full
    }

    #Write-Verbose $Session -Verbose
    #Write-Verbose $Name -Verbose

    $useType = $null

    $useType = [System.Reflection.Assembly]::LoadFile($EntitiesFile)

    if($help -eq $null){return $null}

    if($useType)
    {        
        $h = New-Object -TypeName PowerShellPowered.PowerShellConnect.Entities.PSCommandHelp
        $h.Category = ($help.Category | Out-String -Width 10000).Trim()
        $h.Component = ($help.Component | Out-String -Width 10000).Trim()
        $h.Description = ($help.description | Out-String -Width 10000).Trim()
        $h.Details = ($help.details | Out-String -Width 10000).Trim()
        $h.Examples = ($help.examples | Out-String -Width 10000).Trim()
        $h.FullHelpText = ($help | Out-String -Width 10000).Trim()
        $h.Functionality = ($help.Functionality | Out-String -Width 10000).Trim()
        $h.InputTypes = ($help.inputTypes | Out-String -Width 10000).Trim()
        $h.Name = ($help.Name | Out-String -Width 10000).Trim()
        $h.NonTerminatingErrors = ($help.nonTerminatingErrors | Out-String -Width 10000).Trim()
        $h.Parameters = ($help.parameters | Out-String -Width 10000).Trim()
        GetNavigationLinks -help $help | %{$h.RelatedLinks.Add($_)}
        #$h.RelatedLinks.Add((GetNavigationLinks -help $help).length)
        $h.ReturnValues = ($help.returnValues | Out-String -Width 10000).Trim()
        $h.Role = ($help.Role | Out-String -Width 10000).Trim()
        $h.Synopsis = ($help.Synopsis | Out-String -Width 10000).Trim()
        $h.Syntax = ($help.syntax | Out-String -Width 10000).Trim()
        $h.TerminatingErrors = ($help.terminatingErrors | Out-String -Width 10000).Trim() 

        #$h.PSCmdletHelpExamples = 
        foreach($e in $help.examples.example)
        {
            $ex = New-Object -TypeName PowerShellPowered.PowerShellConnect.Entities.PSCommandHelpExample
            $ex.Code = ($e.code | Out-String -Width 10000).Trim()
            $ex.CommandLines = ($e.commandLines.commandLine.commandText | Out-String -Width 10000).Trim()
            $ex.Introduction = ($e.introduction | Out-String -Width 10000).Trim()
            $ex.Remarks = ($e.remarks | Out-String -Width 10000).Trim()
            $ex.Title = ($e.title | Out-String -Width 10000).Trim()
            $ex.ExampleFullText = ($e | Out-String -Width 10000).Trim()
            
            $h.PSCommandHelpExamples.Add($ex)
        }
        
        #$h.PSCmdletHelpParameters = 
        foreach($p in $help.parameters.parameter)
        {
            #Write-Verbose $p.name -Verbose
            $pi = New-Object -TypeName PowerShellPowered.PowerShellConnect.Entities.PSCommandHelpParameter
            $pi.DefaultValue = ($p.defaultValue | Out-String  -Width 10000).Trim()
            $pi.Description = ($p.description | Out-String  -Width 10000).Trim()
            $pi.Globbing = $p.globbing
            $pi.Name = ($p.name | Out-String  -Width 10000).Trim()
            $pi.ParameterValue = ($p.parameterValue | Out-String  -Width 10000).Trim()
            $pi.PipelineInput = ParseBool -str  $p.pipelineInput -indexOfTrue $true -nullable $true
            $pi.Position = $p.position
            $pi.Required = $p.required.ToBoolean($null)
            $pi.Type = ($p.type.name | Out-String  -Width 10000).Trim()
            $pi.VariableLength = $p.variableLength.ToBoolean($null)
            $h.PSCommandHelpParameters.Add($pi)
        }

        #$h.PSCmdletHelpSyntaxes = 
        foreach($s in $help.syntax.syntaxItem)
        {
            $sx = New-Object -TypeName PowerShellPowered.PowerShellConnect.Entities.PSCommandHelpSyntax
            $sx.Name = ($s.name | Out-String -Width 10000).Trim()
            #$sx.Syntax
            $i=0;
            foreach($sp in $s.parameter)
            {
                $sxi = New-Object -TypeName PowerShellPowered.PowerShellConnect.Entities.PSCommandHelpSyntaxParameter
                $sxi.Name = $sp.name
                $sxi.ParameterValue = $sp.parameterValue                
                $sxi.PipelineInput = ParseBool -str $sp.pipelineInput -indexOfTrue $true -nullable $true
                #$sxi.PipelineInputType = $sp.                
                $sxi.Position = $sp.position
                $sxi.Required = $sp.required.ToBoolean($null)
                
                $sx.PSCommandHelpSyntaxParameters.Add($sxi)                
            }
            $h.PSCommandHelpSyntaxes.Add($sx)
        }
    }
    else
    {        
        $h = New-Object -TypeName psobject
        $h | Add-Member -MemberType NoteProperty -Name alertSet -Value ($help.alertSet | Out-String -Width 10000).Trim()
        $h | Add-Member -MemberType NoteProperty -Name Category -Value ($help.Category | Out-String -Width 10000).Trim()
        $h | Add-Member -MemberType NoteProperty -Name Component -Value ($help.Component | Out-String -Width 10000).Trim()
        $h | Add-Member -MemberType NoteProperty -Name description -Value ($help.description | Out-String -Width 10000).Trim()
        $h | Add-Member -MemberType NoteProperty -Name details -Value ($help.details | Out-String -Width 10000).Trim()
        $h | Add-Member -MemberType NoteProperty -Name examples -Value ($help.examples | Out-String -Width 10000).Trim()
        $h | Add-Member -MemberType NoteProperty -Name Functionality -Value ($help.Functionality | Out-String -Width 10000).Trim()
        $h | Add-Member -MemberType NoteProperty -Name inputTypes -Value ($help.inputTypes | Out-String -Width 10000).Trim()
        $h | Add-Member -MemberType NoteProperty -Name ModuleName -Value ($help.ModuleName | Out-String -Width 10000).Trim()
        $h | Add-Member -MemberType NoteProperty -Name Name -Value ($help.Name | Out-String -Width 10000).Trim()
        $h | Add-Member -MemberType NoteProperty -Name nonTerminatingErrors -Value ($help.nonTerminatingErrors | Out-String -Width 10000).Trim()
        $h | Add-Member -MemberType NoteProperty -Name parameters -Value ($help.parameters | Out-String -Width 10000).Trim()
        $h | Add-Member -MemberType NoteProperty -Name PSSnapIn -Value ($help.PSSnapIn | Out-String -Width 10000).Trim()
        $h | Add-Member -MemberType NoteProperty -Name relatedLinks -Value ($help.relatedLinks | Out-String -Width 10000).Trim()
        $h | Add-Member -MemberType NoteProperty -Name returnValues -Value ($help.returnValues | Out-String -Width 10000).Trim()
        $h | Add-Member -MemberType NoteProperty -Name Role -Value ($help.Role | Out-String -Width 10000).Trim()
        $h | Add-Member -MemberType NoteProperty -Name Synopsis -Value ($help.Synopsis | Out-String -Width 10000).Trim()
        $h | Add-Member -MemberType NoteProperty -Name syntax -Value ($help.syntax | Out-String -Width 10000).Trim()
        $h | Add-Member -MemberType NoteProperty -Name terminatingErrors -Value ($help.terminatingErrors | Out-String -Width 10000).Trim()
        $h | Add-Member -MemberType NoteProperty -Name FullHelpText -Value ($help | Out-String -Width 10000).Trim()

        [PSObject[]]$params = @()
        foreach($p in $help.parameters.parameter)
        {
            $pi = New-Object -TypeName PSObject
            $ti = Get-Member -InputObject $p -MemberType Properties
            foreach($t in $ti){        
                $pi | Add-Member -MemberType NoteProperty -Name ($t.Name | Out-String).Trim() -Value (getparameterdetail $t $p | out-string -Width 10000).trim()
            }
            $params +=$pi
        }
        #$px = New-Object -TypeName PSObject
        #$px | Add-Member -MemberType NoteProperty -Name parameter -Value $params
        #$h | Add-Member -MemberType NoteProperty -Name ParametersEx -Value $px
        $h | Add-Member -MemberType NoteProperty -Name ParametersEx -Value $params

        [PSObject[]]$examples = @()
        foreach($e in $help.examples.example)
        {
            $ex = New-Object -TypeName PSObject
            $ex | Add-Member -MemberType NoteProperty -Name code -Value ($e.code | Out-String -Width 10000).Trim()
            $ex | Add-Member -MemberType NoteProperty -Name commandLines -Value ($e.commandLines.commandLine.commandText | Out-String -Width 10000).Trim()
            $ex | Add-Member -MemberType NoteProperty -Name introduction -Value ($e.introduction | Out-String -Width 10000).Trim()
            $ex | Add-Member -MemberType NoteProperty -Name remarks -Value ($e.remarks | Out-String -Width 10000).Trim()
            $ex | Add-Member -MemberType NoteProperty -Name title -Value ($e.title | Out-String -Width 10000).Trim()
            $examples +=$ex
        }
        $h | Add-Member -MemberType NoteProperty -Name ExamplesEx -Value $examples

        [PSObject[]]$syntaxes = @()

        foreach($s in $help.syntax.syntaxItem)
        {
            $sx = New-Object -TypeName PSObject
            $sx | Add-Member -MemberType NoteProperty -Name name -Value ($s.name | Out-String -Width 10000).Trim()

            [PSObject[]]$sxis = @()       

            $i=0;
            foreach($sp in $s.parameter)
            {
                $sxi = New-Object -TypeName PSObject
                $sxi | Add-Member -MemberType NoteProperty -Name aliases -Value ($sp.aliases | Out-String -Width 10000).Trim()
                $sxi | Add-Member -MemberType NoteProperty -Name description -Value ($sp.description | Out-String -Width 10000).Trim()
                $sxi | Add-Member -MemberType NoteProperty -Name globbing -Value ($sp.globbing | Out-String -Width 10000).Trim()
                $sxi | Add-Member -MemberType NoteProperty -Name name -Value ($sp.name | Out-String -Width 10000).Trim()
                $sxi | Add-Member -MemberType NoteProperty -Name parameterValue -Value ($sp.parameterValue | Out-String -Width 10000).Trim()
                $sxi | Add-Member -MemberType NoteProperty -Name pipelineInput -Value ($sp.pipelineInput | Out-String -Width 10000).Trim()
                $sxi | Add-Member -MemberType NoteProperty -Name position -Value ($sp.position | Out-String -Width 10000).Trim()
                $sxi | Add-Member -MemberType NoteProperty -Name required -Value ($sp.required | Out-String -Width 10000).Trim()
                $sxi | Add-Member -MemberType NoteProperty -Name variableLength -Value ($sp.variableLength | Out-String -Width 10000).Trim()
                $sxis+=$sxi
            }
            $sx | Add-Member -MemberType NoteProperty -Name parameters -Value $sxis        
            $syntaxes +=$sx
        }
        $h | Add-Member -MemberType NoteProperty -Name SyntaxEx -Value $syntaxes
    }

    return $h
";

            //!++ code for win7/psv2, comment on win8 - win8/psv3 has bug for get-help -full, using regular to find more need to change after bug is fixed
            internal const string NameScript = "(Get-Help {0} -Full).Name";
            internal const string SynopsisScript = "(Get-Help {0} -Full).Synopsis";
            internal const string Detailscript = "(Get-Help {0} -Full).details";
            internal const string SyntaxScript = "(Get-Help {0} -Full).syntax";
            internal const string DescriptionScript = "(Get-Help {0} -Full).description ";
            internal const string ParametersScript = "(Get-Help {0} -Full).parameters";
            internal const string InputtypesScript = "(Get-Help {0} -Full).inputTypes";
            internal const string ReturnValuesScript = "(Get-Help {0} -Full).returnValues";
            internal const string TerminatingErrorsScript = "(Get-Help {0} -Full).terminatingErrors";
            internal const string NonTerminatingErrorsScript = "(Get-Help {0} -Full).nonTerminatingErrors";
            internal const string ExamplesScript = "(Get-Help {0} -Full).examples";
            internal const string RelatedLinksScript = "(Get-Help {0} -Full).relatedLinks";
            internal const string CategoryScript = "(Get-Help {0} -Full).Category";
            internal const string ComponentScript = "(Get-Help {0} -Full).Component";
            internal const string RoleScript = "(Get-Help {0} -Full).Role";
            internal const string FunctionalityScript = "(Get-Help {0} -Full).Functionality";
            internal const string FullHelpScript = "Get-Help {0} -Full";


            //!++ code for win8/psv3, uncomment on win8 - win8/psv3 has bug for get-help -full, using regular to find more need to change after bug is fixed
            //internal const string NameScript = "(Get-Help {0}).Name";
            //internal const string SynopsisScript = "(Get-Help {0}).Synopsis";
            //internal const string Detailscript = "(Get-Help {0}).details";
            //internal const string SyntaxScript = "(Get-Help {0}).syntax";
            //internal const string DescriptionScript = "(Get-Help {0}).description ";
            //internal const string ParametersScript = "(Get-Help {0}).parameters";
            //internal const string InputtypesScript = "(Get-Help {0}).inputTypes";
            //internal const string ReturnValuesScript = "(Get-Help {0}).returnValues";
            //internal const string TerminatingErrorsScript = "(Get-Help {0}).terminatingErrors";
            //internal const string NonTerminatingErrorsScript = "(Get-Help {0}).nonTerminatingErrors";
            //internal const string ExamplesScript = "(Get-Help {0}).examples";
            //internal const string RelatedLinksScript = "(Get-Help {0}).relatedLinks";
            //internal const string CategoryScript = "(Get-Help {0}).Category";
            //internal const string ComponentScript = "(Get-Help {0}).Component";
            //internal const string RoleScript = "(Get-Help {0}).Role";
            //internal const string FunctionalityScript = "(Get-Help {0}).Functionality";
            //internal const string FullHelpScript = "Get-Help {0}";

            /*
                (Get-Help get-mailbox -full).details
                Write-Host SYNTAX
                (Get-Help get-mailbox -full).syntax | Out-String
                Write-Host DESCRIPTION
                (Get-Help get-mailbox -full).DESCRIPTION | Out-String
                Write-Host PARAMETERS
                (Get-Help get-mailbox -full).PARAMETERS | Out-String
                Write-Host INPUTS
                (Get-Help get-mailbox -full).INPUTS | Out-String
                Write-Host OUTPUTS
                (Get-Help get-mailbox -full).OUTPUTS | Out-String
                Write-Host TERMINATING ERRORS
                (Get-Help get-mailbox -full).TERMINATINGERRORS | Out-String
                Write-Host NON-TERMINATING ERRORS
                (Get-Help get-mailbox -full).NONTERMINATINGERRORS | Out-String

                (Get-Help get-mailbox -full).examples | Out-String
                Write-Host RELATED LINKS
                (Get-Help get-mailbox -full).RELATEDLINKS | Out-String
                */


        }

        public static class PSVariableNameStrings
        {

            internal const string Credential = "cred";

            internal const string Session = "PSSession";

        }

        public static class PSParameterNameStrings
        {

            internal const string Name = "Name";

            internal const string Value = "Value";

            internal const string Session = "Session";

            internal const string CommandType = "CommandType";
            public const string Module = "Module";


        }

        public static class PSParameterValueString
        {

            internal const string AllPowerShellNative = "Cmdlet,Function,Script,ExternalScript";
        }
    }
}
