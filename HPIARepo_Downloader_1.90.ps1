<#
    HP Image Assistant and Softpaq Repository Downloader
    by Dan Felman/HP Technical Consultant
        
        Loop Code: The 'HPModelsTable' loop code, is in a separate INI.ps1 file 
            ... created by Gary Blok's (@gwblok) post on garytown.com.
            ... https://garytown.com/create-hp-bios-repository-using-powershell
 
        Logging: The Log function based on code by Ryan Ephgrave (@ephingposh)
            ... https://www.ephingadmin.com/powershell-cmtrace-log-function/

    Version 1.00 - Initial code release
    Version 1.10
        added ability to maintain single repository - setting also kept int INI file
        added function to try and import and load HPCMSL powershell module
    Version 1.11
        fix GUI when selecting/deslection All rows checkmark (added call to get_filters to reset categories)
        added form button (at the bottom of the form) to clear textbox
    Version 1.12
        fix interface issues when selecting/deselecting all rows (via header checkmark)
    Versoin 1.15
        added ability for script to use a single common repository, with new variable in INI file
        moved $DebugMode variable to INI file
    Version 1.16
        fixed UI issue switching from single to non-single repository
        Added color coding of which path is in use (single or multiple share repo paths)
    Version 1.20
        Fixed single repository filter cleanup... was not removing all other platform filters previously
        Added function to show Softpaqs added (or not) after a Sync
        Added button to reread model filters from repositories and refresh the grid
    Version 1.25
        Code cleanup of Function sync_individual_repositories 
        changed 'use single repository' checkbox to radio buttons on path fields
        Added 'Distribute SCCM Packages' ($v_DistributeCMPackages) variable use in INI file
            -- when selected, sends command to CM to Distribute packages
    Version 1.30
        Added ability to sync specific softpaqs by name - listed in INI file
            -- added SqName entry to $HPModelsTable list to hold special softpaqs needed/model
    Version 1.31
        moved Debug Mode checkbox to bottom
        added HPIA folder view, and HPIA button to initiate a create/update package in CM
    Version 1.32
        Moved HPIA info to SCCM block in UI
    Version 1.40
        increased windows size based on feedback
        added checkmark to keep existing category filters - useful when maintaining Softpaqs for more than a single OS Version
        added checks for, and report on, CMSL Repository Sync errors 
        added function to list current category filters
        added separate function to modify setting in INI.ps1 file
        added buttons to increase and decrease output textbox text size
    Version 1.41
        added IP lookup for internet connetion local and remote - useful for debugging... posted to HPA's log file
    Version 1.45
        added New Function w/job to trace connections to/from HP (not always 100% !!!)
        added log file name path to UI
    Version 1.50
        added protection against Platform/OS Version not supported
        Added ability to start a New log file w/checkmark in IU - log saved to log.bak0,log.bak1, etc.
    Version 1.51
        Removed All column (was selecting all categories with single click
        Added 'SqName' column to allow/disallow software listed in INI.ps1 file from being downloaded
        Fixed UI issues when clicking/selecting or deselcting categories and Model row checkbox
    Version 1.60
        developed non-UI runtime options to support using script on a schedule
           - added runtime options -Help, -IniFile <option>, -RepoStyle <option>, -Products <list>, -ListFilters -Sync, -noIniSw, -ShowActivityLog, -newLog
           - added function list_filters
           - added function sync_repos 
    Versionj 1.61/2 - scriptable testing version
    Version 1.65
        added support for INI Software named by ID, not just by name
        added Browse button to find HPIA folder
    Version 1.66
        add [Sync] dialog to ask if unselected products' filters should be removed from common repository
    Version 1.70
        add setting checkbox to keep going on Sync missing file errors
    Version 1.75
        add platform and OS version check, to validate version is supported for selected platforms
        -- advisory mode only... User will be informed in the OS version selected is NOT supported
        -- by a platform selected in the list
        renamed Done to Exit button
        added ability to Import an existing HPIA sync'd Repository
    Version 1.80-3
        added repositories Browse buttons for both Common (shared) and Individual (rooted) folders
        cleaned up code, fixes issues found
        improved fixture to maintain additional software softpaqs outside of CMSL sync command
    Version 1.85
        added ability to Add models to list
    Version 1.86/7
        added Addons entries support for when adding models, and reporting of what is added
    Version 1.88
        Fixes
    Version 1.90
        Complete rewrite of 'AddOns' functionality. No longer a list of softpaqs in INI.ps1
        Now uses a Platform ID as a flag file in the .ADDSOFTWARE folder
        Requires update to INI file in terms of the HPModelsTable, as .Addons no longer used in script
#>
<#
    Args
        -inifile - full path to ini.ps1 file , also 1st arg in command line
#>
[CmdletBinding()]
param(
    [Parameter( Mandatory = $false, Position = 0, ValueFromRemainingArguments )]
    [Switch]$Help,                                # $help is set to $true if '-help' passed as argument
    [Parameter( Mandatory = $false, Position = 1, HelpMessage="Path to ini.ps1 file. Default: 'HPIARepo_ini.ps1'" )]
    [AllowEmptyString()]
    [string]$inifile=".\HPIARepo_ini.ps1",
    [Parameter( Mandatory = $false, Position = 2 )]
    [ValidateSet('individual', 'Common')]
    [string]$RepoStyle='individual',
    [Parameter( Mandatory = $false, Position = 3 )]
    [array]$products=$null,
    [Parameter( Mandatory = $false )]
    [Switch]$ListFilters,
    [Parameter( Mandatory = $false )]
    [Switch]$noIniSw=$false,
    [Parameter( Mandatory = $false )]
    [Switch]$showActivityLog=$false,
    [Parameter( Mandatory = $false )]
    [Switch]$newLog=$false,
    [Parameter( Mandatory = $false, Position = 0, ValueFromRemainingArguments )]
    [Switch]$Sync                                 # $help is set to $true if '-help' passed as argument                                            
) # param

$ScriptVersion = "1.90.02 (October 31, 2021)"

Function Show_Help {
    
} # Function Show_Help

#=====================================================================================
# manage script runtime parameters
#=====================================================================================

if ( $help ) {
    "`nRunning script without parameters opens in UI mode. This mode should be used to set up needed filters and develop Repositories"
    "`nRuntime options:"
    "`n[-help|-h] will display this text"
    "`nHPIARepo_Downloader_1.90.ps1 [[-Sync] [-ListFilters] [-inifile .\<filepath>\HPIARepo_ini.ps1] [-RepoStyle common|individual] [-products '80D4,8549,8470'] [-NoIniSw] [-ShowActivityLog]]`n"
    "... <-IniFile>, <-RepoStyle>, <-Productsts> parameters can also be positional without parameter names. In this case, every parameter counts"    
    "`nExample: HPIARepo_Downloader_1.90.ps1 .\<path>\HPIARepo_ini.ps1 Common '80D4,8549,8470'`n"    
    "-ListFilters`n`tlist repository filters in place on selected product repositories"
    "-IniFile <path to INI.ps1>`n`tthis option can be used when running from a script to set up different downloads."
    "`tIf option not give, it will default to .\HPIARepo_ini.ps1"
    "-RepoStyle {Common|Individual}`n`tthis option selects the repository style used by the downloader script"
    "`t`t'Common' - There will be a single repository used for all models - path extracted from INI.ps1 file"
    "`t`t'Individual' - Each model will have its own repository folder"
    "-Products '1234', '2222'`n`ta list of HP Model Product codes, as example '80D4,8549,8470'"
    "`tIf omitted, any entry in the INI.ps1 file that has a repository created will be updated"
    "-NoIniSw`n`tPrevent from syncing Softpaqs listed by name in INI file"
    "-showActivityLog`n`tShow output from CMSL Sync/Cleanup activity log"
    "`tThis option is useful when using -Sync"
    "-newLog`n`tstart new log file in script's directory - backup current log"
    "-Sync`n`tUse Hp CMSL commands to sync repositories. Command assumes repositories already created with script in UI mode"
    exit
}

$RunUI = $true                                        # default to displaying a UI

#=====================================================================================

# get the path to the running script, and populate name of INI configuration file
$scriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

Set-Location $ScriptPath
'Script invoked from: '+$ScriptPath

#--------------------------------------------------------------------------------------
#$IniFile = "HPIARepo_ini.ps1"                        # assume this INF file in same location as script

#$IniFileRooted = [System.IO.Path]::IsPathRooted($IniFile)

if ( [System.IO.Path]::IsPathRooted($IniFile) ) {
    $IniFIleFullPath = $IniFile
} else {
    $IniFIleFullPath = "$($ScriptPath)\$($IniFile)"
} # else if ( [System.IO.Path]::IsPathRooted($IniFile) )

if ( -not (Test-Path $IniFIleFullPath) ) {
    "-IniFile '$IniFIleFullPath' file not found!!! Can't continue"
    exit
}

. $IniFIleFullPath                                   # source/read the code in the INI file      

#--------------------------------------------------------------------------------------
# add required .NET framwork items to support GUI
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

#--------------------------------------------------------------------------------------
# Script Environment Vars

$CMConnected = $false                                # is a connection to SCCM established?
$SiteCode = $null

$s_ModelAdded = $False                                 # set $True when user adds models to list in the UI
$s_AddSoftware = '.ADDSOFTWARE'                        # sub-folders where named downloaded Softpaqs will reside
$HPIAActivityLog = 'activity.log'                    # name of HPIA activity log file
$HPServerIPFile = "$($ScriptPath)\15"                # used to temporarily hold IP connections, while job is running

$MyPublicIP = (Invoke-WebRequest ifconfig.me/ip).Content.Trim()
'My Public IP: '+$MyPublicIP | Out-Host

#--------------------------------------------------------------------------------------

$TypeError = -1
$TypeNorm = 1
$TypeWarn = 2
$TypeDebug = 4
$TypeSuccess = 5
$TypeNoNewline = 10

#=====================================================================================
#region: CMTraceLog Function formats logging in CMTrace style
function CMTraceLog {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $false)] $Message,
		[Parameter(Mandatory = $false)] $ErrorMessage,
		[Parameter(Mandatory = $false)] $Component = "HP HPIA Repository Downloader",
		[Parameter(Mandatory = $false)] [int]$Type
	)
	<#
    Type: 1 = Normal, 2 = Warning (yellow), 3 = Error (red)
    #>
	$Time = Get-Date -Format "HH:mm:ss.ffffff"
	$Date = Get-Date -Format "MM-dd-yyyy"

	if ($null -ne $ErrorMessage) { $Type = $TypeError }
	if ($Component -eq $null) { $Component = " " }
	if ($null -eq $Type) { $Type = $TypeNorm }

	$LogMessage = "<![LOG[$Message $ErrorMessage" + "]LOG]!><time=`"$Time`" date=`"$Date`" component=`"$Component`" context=`"`" type=`"$Type`" thread=`"`" file=`"`">"

    #$Type = 4: Debug output ($TypeDebug)
    #$Type = 10: no \newline ($TypeNoNewline)

    if ( ($Type -ne $TypeDebug) -or ( ($Type -eq $TypeDebug) -and $DebugMode) ) {
        $LogMessage | Out-File -Append -Encoding UTF8 -FilePath $Script:v_LogFile

        # Add output to GUI message box
        OutToForm $Message $Type $Script:TextBox
        
    } else {
        $lineNum = ((get-pscallstack)[0].Location -split " line ")[1]    # output: CM_HPIARepo_Downloader.ps1: line 557
        Write-Host "$lineNum $(Get-Date -Format "HH:mm:ss") - $Message"
    }
} # function CMTraceLog

#=====================================================================================
<#
    Function OutToForm
        Designed to output message to Form's message box
        it uses color coding of text for different message types
#>
Function OutToForm { 
	[CmdletBinding()]
	param( $pMessage, [int]$pmsgType, $pTextBox)

    if ( $pmsgType -eq $TypeDebug ) {
        $pMessage = '{dbg}'+$pMessage
    }
    if ($RunUI) { 
        switch ( $pmsgType ) {
       -1 { $pTextBox.SelectionColor = "Red" }                  # Error
        1 { $pTextBox.SelectionColor = "Black" }                # default color is black
        2 { $pTextBox.SelectionColor = "Brown" }                # Warning
        4 { $pTextBox.SelectionColor = "Orange" }               # Debug Output
        5 { $pTextBox.SelectionColor = "Green" }                # success details
        10 { $pTextBox.SelectionColor = "Black" }               # do NOT add \newline to message output
        } # switch ( $pmsgType )

        # message Tpye = 10/$TypeNeNewline prevents a nl so next output is written contiguous

        if ( $pmsgType -eq $TypeNoNewline ) {
            $pTextBox.AppendText("$($pMessage) ")
        } else {
            $pTextBox.AppendText("$($pMessage) `n")
        }
        $pTextBox.Refresh()
        $pTextBox.ScrollToCaret()
    } else {
        $pMessage | Out-Host
    }
} # Function OutToForm

#=====================================================================================
<#
    Function Load_HPCMSLModule
        The function will test if the HP Client Management Script Library is loaded
        and attempt to load it, if possible
#>
function Load_HPCMSLModule {

    if ( $DebugMode ) { CMTraceLog -Message "> Load_HPCMSLModule" -Type $TypeNorm }
    $m = 'HPCMSL'

    CMTraceLog -Message "Checking for required HP CMSL modules... " -Type $TypeNoNewline

    # If module is imported say that and do nothing
    if (Get-Module | Where-Object {$_.Name -eq $m}) {
        if ( $DebugMode ) { write-host "Module $m is already imported." }
        CMTraceLog -Message "Module already imported." -Type $TypSuccess
    } else {

        # If module is not imported, but available on disk then import
        if (Get-Module -ListAvailable | Where-Object {$_.Name -eq $m}) {
            if ( $DebugMode ) { write-host "Importing Module $m." }
            CMTraceLog -Message "Importing Module $m." -Type $TypeNoNewline
            Import-Module $m -Verbose
            CMTraceLog -Message "Done" -Type $TypSuccess
        } else {

            # If module is not imported, not available on disk, but is in online gallery then install and import
            if (Find-Module -Name $m | Where-Object {$_.Name -eq $m}) {
                if ( $DebugMode ) { write-host "Upgrading NuGet and updating PowerShellGet first." }
                CMTraceLog -Message "Upgrading NuGet and updating PowerShellGet first." -Type $TypSuccess
                if ( !(Get-packageprovider -Name nuget -Force) ) {
                    Install-PackageProvider -Name NuGet -ForceBootstrap
                }
                # next is PowerShellGet
                $lPSGet = find-module powershellget
                write-host 'Installing PowerShellGet version' $lPSGet.Version
                Install-Module -Name PowerShellGet -Force # -Verbose

                Write-Host 'should restart PowerShell after upating module PoweShellGet'

                # and finally module HPCMSL
                if ( $DebugMode ) { write-host "Installing and Importing Module $m." }
                CMTraceLog -Message "Installing and Importing Module $m." -Type $TypSuccess
                Install-Module -Name $m -Force -SkipPublisherCheck -AcceptLicense -Scope CurrentUser #  -Verbose 
                Import-Module $m -Verbose
                CMTraceLog -Message "Done" -Type $TypSuccess
            } else {

                # If module is not imported, not available and not in online gallery then abort
                write-host "Module $m not imported, not available and not in online gallery, exiting."
                CMTraceLog -Message "Module $m not imported, not available and not in online gallery, exiting." -Type $TypError
                exit 1
            }
        } # else if (Get-Module -ListAvailable | Where-Object {$_.Name -eq $m}) 
    } # else if (Get-Module | Where-Object {$_.Name -eq $m})

    if ( $DebugMode ) { CMTraceLog -Message "< Load_HPCMSLModule" -Type $TypeNorm }

} # function Load_HPCMSLModule

#=====================================================================================
<#
    Function Test_CMConnection
        The function will test the CM server connection
        and that the Task Sequences required for use of the Script are available in CM
        - will also test that both download and share paths exist
#>
Function Test_CMConnection {

    if ( $DebugMode ) { CMTraceLog -Message "> Test_CMConnection" -Type $TypeNorm }

    if ( $Script:CMConnected ) { return $True }                  # already Tested connection

    $pCurrentLoc = Get-Location

    CMTraceLog -Message "Connecting to CM Server: ""$FileServerName""" -Type $TypeNoNewline
    
    #--------------------------------------------------------------------------------------
    # check for ConfigMan  on this server, and source the PS module

    $boolConnectionRet = $False

    if (Test-Path $env:SMS_ADMIN_UI_PATH) {
        $lCMInstall = Split-Path $env:SMS_ADMIN_UI_PATH
        Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1)

        #--------------------------------------------------------------------------------------
        # by now we know the script is running on a CM server and the PS module is loaded
        # so let's get the CMSite content info
    
        Try {
            $Script:SiteCode = (Get-PSDrive -PSProvider CMSite).Name               # assume CM PS modules loaded at this time

            if (Test-Path $lCMInstall) {        
                try { Test-Connection -ComputerName "$FileServerName" -Quiet
                    CMTraceLog -Message " ...Connected" -Type $TypeSuccess 
                    $boolConnectionRet = $True
                    $CMGroupAll.Text = 'SCCM - Connected'
                }
                catch {
	                CMTraceLog -Message "Not Connected to File Server, Exiting" -Type $TypeError 
                }
            } else {
                CMTraceLog -Message "CM Installation path NOT FOUND: '$lCMInstall'" -Type $TypeError 
            } # else
        } # Try
        Catch {
            CMTraceLog -Message "Error obtaining CM's CMSite provider on this server" -Type $TypeError
        } # Catch
    } else {
        CMTraceLog -Message "Can't find CM Installation on this system" -Type $TypeError
    }
    if ( $DebugMode ) { CMTraceLog -Message "< Test_CMConnection" -Type $TypeNorm }

    Set-Location $pCurrentLoc

    return $boolConnectionRet

} # Function Test_CMConnection

#=====================================================================================
<#
    Function CM_HPIAPackage
#>
Function CM_HPIAPackage {
    [CmdletBinding()]
	param( $pHPIAPkgName, $pHPIAPath, $pHPIAVersion )   

    if ( Test_CMConnection ) {
        $pCurrentLoc = Get-Location

        if ( $DebugMode ) { CMTraceLog -Message "... setting location to: $($SiteCode)" -Type $TypeDebug }
        Set-Location -Path "$($SiteCode):"

        #--------------------------------------------------------------------------------
        # now, see if HPIA package exists
        #--------------------------------------------------------------------------------

        $lCMHPIAPackage = Get-CMPackage -Name $pHPIAPkgName -Fast

        if ( $null -eq $lCMHPIAPackage ) {
            if ( $DebugMode ) { CMTraceLog -Message "... HPIA Package missing... Creating New" -Type $TypeNorm }
            $lCMHPIAPackage = New-CMPackage -Name $pHPIAPkgName -Manufacturer "HP"
            CMTraceLog -ErrorMessage "... HPIA package created - PackageID $lCMHPIAPackage.PackageId"
        } else {
            CMTraceLog -Message "... HPIA package found - PackageID $($lCMHPIAPackage.PackageId)"
        }
        #if ( $DebugMode ) { CMTraceLog -Message "... setting HPIA Package to Version: $($v_OSVER), path: $($pRepoPath)" -Type $TypeDebug }

        CMTraceLog -Message "... HPIA package - setting name ''$pHPIAPkgName'', version ''$pHPIAVersion'', Path ''$pHPIAPath''" -Type $TypeNorm
	    Set-CMPackage -Name $pHPIAPkgName -Version $pHPIAVersion
	    Set-CMPackage -Name $pHPIAPkgName -Path $pHPIAPath

        if ( $Script:v_DistributeCMPackages  ) {
            CMTraceLog -Message "... updating CM Distribution Points"
            update-CMDistributionPoint -PackageId $lCMHPIAPackage.PackageID
        }

        $lCMHPIAPackage = Get-CMPackage -Name $pHPIAPkgName -Fast                               # make sure we are woring with updated/distributed package
        CMTraceLog -Message "... HPIA package updated - PackageID $($lCMHPIAPackage.PackageId)" -Type $TypeNorm

        #--------------------------------------------------------------------------------
        Set-Location -Path $pCurrentLoc 
    } else {
        CMTraceLog -ErrorMessage "NO CM Connection to update HPIA package"
    }
    CMTraceLog -Message "< HCM_HPIAPackage DONE" -Type $TypeSuccess
                           
} # CM_HPIAPackage

#=====================================================================================
<#
    Function CM_RepoUpdate
#>
Function CM_RepoUpdate {
    [CmdletBinding()]
	param( $pModelName, $pModelProdId, $pRepoPath )                             

    $pCurrentLoc = Get-Location

    if ( $DebugMode ) {  CMTraceLog -Message "> CM_RepoUpdate" -Type $TypeNorm }

    # develop the Package name
    $lPkgName = 'HP-'+$pModelProdId+'-'+$pModelName
    CMTraceLog -Message "... updating repository for SCCM package: $($lPkgName)" -Type $TypeNorm

    if ( $DebugMode ) { CMTraceLog -Message "... setting location to: $($SiteCode)" -Type $TypeDebug }
    Set-Location -Path "$($SiteCode):"

    if ( $DebugMode ) { CMTraceLog -Message "... getting CM package: $($lPkgName)" -Type $TypeDebug }
    $lCMRepoPackage = Get-CMPackage -Name $lPkgName -Fast

    if ( $null -eq $lCMRepoPackage ) {
        CMTraceLog -Message "... Package missing... Creating New" -Type $TypeNorm
        $lCMRepoPackage = New-CMPackage -Name $lPkgName -Manufacturer "HP"
    }
    #--------------------------------------------------------------------------------
    # update package with info from share folder
    #--------------------------------------------------------------------------------
    if ( $DebugMode ) { CMTraceLog -Message "... setting CM Package to Version: $($v_OSVER), path: $($pRepoPath)" -Type $TypeDebug }

	Set-CMPackage -Name $lPkgName -Version "$($v_OSVER)"
	Set-CMPackage -Name $lPkgName -Path $pRepoPath

    if ( $Script:v_DistributeCMPackages  ) {
        CMTraceLog -Message "... updating CM Distribution Points"
        update-CMDistributionPoint -PackageId $lCMRepoPackage.PackageID
    }
    $lCMRepoPackage = Get-CMPackage -Name $lPkgName -Fast                               # make sure we are woring with updated/distributed package

    #--------------------------------------------------------------------------------

    Set-Location -Path $pCurrentLoc

    if ( $DebugMode ) { CMTraceLog -Message "< CM_RepoUpdate" -Type $TypeNorm }

} # Function CM_RepoUpdate

#=====================================================================================
<#
    Function Get_ActivityLogEntries
        obtains log entries from the ..\.Repository\activity.log file
        ... searches for last sync data
#>
Function Get_ActivityLogEntries {
    [CmdletBinding()]
	param( $pRepoFolder )

    if ( (-not $RunUI) -and ($showActivityLog -eq $false) ) {
        return
    }
    $ldotRepository = "$($pRepoFolder)\.Repository"
    $lLastSync = $null
    $lCurrRepoLine = 0
    $lLastSyncLine = 0

    if ( Test-Path $lDotRepository ) {

        #--------------------------------------------------------------------------------
        # find the last 'Sync started' entry 
        #--------------------------------------------------------------------------------
        $lRepoLogFile = "$($ldotRepository)\$($HPIAActivityLog)"

        if ( Test-Path $lRepoLogFile ) {
            $find = 'sync has started'                     # look for this string in HPIA's log
            (Get-Content $lRepoLogFile) | 
                Foreach-Object { 
                    $lCurrRepoLine++ 
                    if ($_ -match $find) { $lLastSync = $_ ; $lLastSyncLine = $lCurrRepoLine } 
                } # Foreach-Object
            CMTraceLog -Message "      [activity.log - last sync @line $lLastSyncLine] $lLastSync " -Type $TypeWarn
        } # if ( Test-Path $lRepoLogFile )

        #--------------------------------------------------------------------------------
        # now, $lLastSyncLine holds the log's line # where the last sync started
        #--------------------------------------------------------------------------------
        if ( $lLastSync ) {

            $lLogFile = Get-Content $lRepoLogFile

            for ( $i = 0; $i -lt $lLogFile.Count; $i++ ) {
                if ( $i -ge ($lLastSyncLine-1) ) {
                    if ( ($lLogFile[$i] -match 'done downloading exe') -or 
                        ($lLogFile[$i] -match 'already exists') ) {

                        CMTraceLog -Message "      [activity.log - Softpaq update] $($lLogFile[$i])" -Type $TypeWarn
                    }
                    if ( ($lLogFile[$i] -match 'Repository Synchronization failed') -or 
                        ($lLogFile[$i] -match 'ErrorRecord') -or
                        ($lLogFile[$i] -match 'WebException') ) {

                        CMTraceLog -Message "      [activity.log - Sync update] $($lLogFile[$i])" -Type $TypeError
                    }
                }
            } # for ( $i = 0; $i -lt $lLogFile.Count; $i++ )
        } # if ( $lLastSync )

    } # if ( Test-Path $lDotRepository )

} # Function Get_ActivityLogEntries

#=====================================================================================
<#
    Function Backup_Log
        Rename the file name in argument with .bak[0..99].log extention
        Args:
            $pLogFileFullPath: file to rename (as backup)
#>
Function Backup_Log {
    [CmdletBinding()]
	param( $pLogFileFullPath )

    if ( Test-Path $pLogFileFullPath ) {
        $lLogFileNameBase = [IO.Path]::GetFileNameWithoutExtension($pLogFileFullPath)
        $lLogFileNameExt = [System.IO.Path]::GetExtension($pLogFileFullPath)

        # log file exists, so back it up/rename it
        for ($i=0; $i -lt 100; $i++ ) {
            $lNewLogFilePath = $ScriptPath+'\'+$lLogFileNameBase+'.bak'+$i.ToString()+$lLogFileNameExt
            if ( -not (Test-Path $lNewLogFilePath) ) {
                CMTraceLog -Message "Renamed existing log file to $($lNewLogFilePath)" -Type $TypeNorm
                $pLogFileFullPath | Rename-Item -NewName $lNewLogFilePath

                return
            } # if ( -not (Test-Path $lNewLogFilePath)
        } # for ($i=0; $i -lt 100; $i++ )
    } # if ( $pLogFileFullPath )

} # Function Backup_Log

#=====================================================================================
<#
    Function init_repository
        This function will create a repository folder
        ... and initialize it for HPIA, if $pInitialize arg is $True
        Args:
            $pRepoFOlder: folder to validate, or create
            $pInitRepository: $true, initialize repository, 
                              $false: do not initialize (used as root of individual repository folders)
#>
Function init_repository {
    [CmdletBinding()]
	param( $pRepoFolder,
            $pInitialize )

    if ( $DebugMode ) { CMTraceLog -Message "> init_repository" -Type $TypeNorm }

    $lCurrentLoc = Get-Location

    $retRepoCreated = $false
    if ( (Test-Path $pRepoFolder) -and ($pInitialize -eq $false) ) {
        $retRepoCreated = $true
    } else {
        $lRepoParentPath = Split-Path -Path $pRepoFolder -Parent        # Get the Parent path to the Main Repo folder

        #--------------------------------------------------------------------------------
        # see if we need to create the path to the folder

        if ( !(Test-Path $lRepoParentPath) ) {
            Try {
                # create the Path to the Repo folder
                New-Item -Path $lRepoParentPath -ItemType directory
                if ( $DebugMode ) { CMTraceLog -Message "Supporting path created: $($lRepoParentPath)" -Type $TypeWarn }
            } Catch {
                if ( $DebugMode ) { CMTraceLog -Message "Supporting path creation Failed: $($lRepoParentPath)" -Type $TypeError }
                CMTraceLog -Message "[init_repository] Done" -Type $TypeNorm
                return $retRepoCreated
            } # Catch
        } # if ( !(Test-Path $lRepoPathSplit) ) 

        #--------------------------------------------------------------------------------
        # now add the Repo folder if it doesn't exist
        if ( !(Test-Path $pRepoFolder) ) {
            CMTraceLog -Message "... creating Repository Folder $pRepoFolder" -Type $TypeNorm
            New-Item -Path $pRepoFolder -ItemType directory
        } # if ( !(Test-Path $pRepoFolder) )

        $retRepoCreated = $true

        #--------------------------------------------------------------------------------
        # if needed, check on repository to initialize (CMSL repositories have a .Repository folder)

        if ( $pInitialize -and !(test-path "$pRepoFolder\.Repository")) {
            Set-Location $pRepoFolder
            $initOut = (Initialize-Repository) 6>&1
            CMTraceLog -Message  "... Repository Initialization done $($Initout)"  -Type $TypeNorm 

            CMTraceLog -Message  "... configuring $($pRepoFolder) for HP Image Assistant" -Type $TypeNorm
            Set-RepositoryConfiguration -setting OfflineCacheMode -cachevalue Enable 6>&1   # configuring the repo for HP IA's use
            # configuring to create 'Contents.CSV' after every Sync
            Set-RepositoryConfiguration -setting RepositoryReport -value csv 6>&1   
            

        } # if ( $pInitialize -and !(test-path "$pRepoFOlder\.Repository"))

    } # else if ( (Test-Path $pRepoFolder) -and ($pInitialize -eq $false) )

    #--------------------------------------------------------------------------------
    # intialize special folder for holding named softpaqs
    # ... this folder will hold softpaqs added outside CMSL' sync/cleanup root folder
    #--------------------------------------------------------------------------------
    $lAddSoftpaqsFolder = "$($pRepoFolder)\$($s_AddSoftware)"

    if ( !(Test-Path $lAddSoftpaqsFolder) -and $pInitialize ) {
        CMTraceLog -Message "... creating Add-on Softpaq Folder $lAddSoftpaqsFolder" -Type $TypeNorm
        New-Item -Path $lAddSoftpaqsFolder -ItemType directory
        if ( $DebugMode ) { CMTraceLog -Message "NO $lAddSoftpaqsFolder" -Type $TypeWarn }
    } # if ( !(Test-Path $lAddSoftpaqsFolder) )

    if ( $DebugMode ) { CMTraceLog -Message "< init_repository" -Type $TypeNorm }

        Set-Location $lCurrentLoc
    return $retRepoCreated

} # Function init_repository

#=====================================================================================
<#
    Function Get_HPServerIPConnection
        This function will run a Sync and a Cleanup commands from HPCMSL

    expects parameter 
        - string of function calling
        - file to store info for further analysis
#>
Function Get_HPServerIPConnection {
    [CmdletBinding()]
	param(  $pJobName,                             # name of job, used in keeping tab of job, and to create specific out file
            $pOutFile,                             # base file name to use for out file
            [int]$pSleep )

    # Get the Calling Function name (for use in temp file output)
    $lCallingFunc = "[$((Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name)]_$($pJobName)"

    # create temp file to hold IP addresses - useful for troubleshooting - use PS command $pJobName in filename
    $temp15File = "$($pOutFile)_$($pJobName).TXT"   

    # if there is a Proxy for Internet access, then it may be impossible to see the actual HP remote IP address
    if ( (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').ProxyEnable -eq 1) {
        CMTraceLog -Message '[Get_HPServerIPConnection] proxy IS enabled - Can''t obtain real HP Server IP' -Type $TypeWarn 
    } else {   
        start-job -Name $pJobName -scriptblock { 
            Start-Sleep $using:pSleep
            foreach ($connection in (Get-NetTCPConnection -appliedsetting internet)) {     # (Get-NetTCPConnection -state established -appliedsetting internet)
                # match Remote IP starting with 15 (HP's class A address)
                if ( $connection.RemoteAddress -match [RegEx]'^15.' ) {
                    "$($using:lCallingFunc) - Remote IP: $($connection.RemoteAddress):$($connection.RemotePort) - My Public IP=$($using:MyPublicIP)" >> $using:temp15File
                }
            } # foreach ($connection in (Get-NetTCPConnection -state established -appliedsetting internet))     
        } # start-job -scriptblock 
    } # else if ( (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').ProxyEnable -eq 1)

} # Get_HPServerIPConnection

#=====================================================================================
<#
    Function Download_AddOn_Softpaqs
        designed to find named Softpaqs from $HPModelsTable and download those into their own repo folder
        Requires: pFolder - repository folder
                  prodcode - SysId of product to check for
#>
Function Download_AddOn_Softpaqs {
[CmdletBinding()] 
    param( $pGrid,
            $pFolder,
            $pCommonRepoFlag )

    if ( $script:noIniSw ) { return }    # $script:noIniSw = runstring option, if executed from command line

    $lCurrentLoc = Get-Location

    $lAddSoftpaqsFolder = "$($pFolder)\$($s_AddSoftware)"   # get location of .ADDSOFTWARE subfolder
    Set-Location $lAddSoftpaqsFolder

    if ( $pCommonRepoFlag ) {
        #--------------------------------------------------------------------------------
        # There are likely multiple models being held in this repository
        # ... so find all Platform IF AddOns flag file for content defined by user
        #--------------------------------------------------------------------------------
        for ( $iRow = 0; $iRow -lt $pGrid.RowCount; $iRow++ ) {
        
            $lProdCode = $pGrid[1,$iRow].Value                      # column 1 = Model/Prod ID
            $lModelName = $pGrid[2,$iRow].Value                     # column 2 = Model name
            $lAddOnsFlag = $pGrid.Rows[$iRow].Cells['AddOns'].Value # column 7 = 'AddOns' checkmark
            $lProdIDFlagFile = $lAddSoftpaqsFolder+'\'+$lProdCode

            if ( $lAddOnsFlag -and (Test-Path $lProdIDFlagFile) ) {

                if ( $DebugMode ) { CMTraceLog -Message 'Flag file found:'+$lAddOnsFlagFile -Type $TypeWarn }
                [array]$lAddOnsList = Get-Content $lProdIDFlagFile

                if ( -not ([String]::IsNullOrWhiteSpace($lAddOnsList)) ) {
                    CMTraceLog -Message "... platform: $($lProdCode): checking AddOns flag file"
                    if ( $DebugMode ) { CMTraceLog -Message 'calling Get-SoftpaqList():'+$lProdCode -Type $TypeNorm }
                    $lSoftpaqList = (Get-SoftpaqList -platform $lProdCode -os $Script:OS -osver $Script:v_OSVER -characteristic SSM) 6>&1  
                    ForEach ( $iEntry in $lAddOnsList ) {
                        $lEntryFound = $False
                        ForEach ( $iSoftpaq in $lSoftpaqList ) {
                            if ( [string]$iSoftpaq.Name -match $iEntry ) {
                                $lEntryFound = $True
                                $lSoftpaqExe = $lAddSoftpaqsFolder+'\'+$iSoftpaq.id+'.exe'
                                if ( Test-Path $lSoftpaqExe ) {
                                    CMTraceLog -Message "... $($iSoftpaq.id) already downloaded - $($iSoftpaq.Name)"
                                } else {
                                    CMTraceLog -Message "`tdownloading $($iSoftpaq.id)/$($iSoftpaq.Name)" -Type $TypeNoNewline
                                    $ret = (Get-Softpaq $iSoftpaq.Id) 6>&1
                                    CMTraceLog -Message "... downloaded" -Type $TypeNorm
                                } # else if ( Test-Path $lSoftpaqExe ) {
                                $ret = (Get-SoftpaqMetadataFile $iSoftpaq.Id) 6>&1   # ALWAYS update the CVA file, even if previously downloaded
                                if ( $DebugMode ) { CMTraceLog -Message  "... Get-SoftpaqMetadataFile done: $($ret)"  -Type $TypeWarn }
                            } # if ( [string]$iSoftpaq.Name -match $iEntry )
                        } # ForEach ( $iSoftpaq in $lSoftpaqList )
                        if ( -not $lEntryFound ) {
                            CMTraceLog -Message  "... '$($iEntry)': Addon softpaq entry not found in softpaq list for this platform and OS version"  -Type $TypeWarn
                        } # if ( -not $lEntryFound )
                    } # ForEach ( $lEntry in $lAddOnsList )

                } else {
                    CMTraceLog -Message $lProdCode': Flag file found but empty: '
                } # else if ( -not ([String]::IsNullOrWhiteSpace($lAddOns)) )

            } else {
                #CMTraceLog -Message "... $($lProdCode): AddOns cell not set, will not attempt to download"
            } # else if (Test-Path $lAddOnsFlagFile)

        } # for ( $i = 0; $i -lt $pGrid.RowCount; $i++ )
    } else {
        #--------------------------------------------------------------------------------
        # Search grid for this model's product ID, so we can find the AddOns flag file
        #--------------------------------------------------------------------------------
        for ( $i = 0; $i -lt $pGrid.RowCount; $i++ ) {
            $lModelName = $pGrid[2,$i].Value                    # column 2 has the Model name
            if ( $lModelName -match (split-path $pFolder -Leaf) ) { # we found the mode, so get the Prod ID
                $lProdCode = $pGrid[1,$i].Value                      # column 1 has the Model/Prod ID
                $lAddOnsFlag = $pGrid.Rows[$i].Cells['AddOns'].Value # column 7 has the AddOns checkmark
            }
        } # for ( $i = 0; $i -lt $pGrid.RowCount; $i++ )

        $lProdIDFlagFile = $lAddSoftpaqsFolder+'\'+$lProdCode

        if ( $lAddOnsFlag -and (Test-Path $lProdIDFlagFile) ) {

            if ( $DebugMode ) { CMTraceLog -Message 'Flag file found:'+$lProdIDFlagFile -Type $TypeWarn }
            [array]$lAddOnsList = Get-Content $lProdIDFlagFile
            if ( -not ([String]::IsNullOrWhiteSpace($lAddOnsList)) ) {
                CMTraceLog -Message "... found AddOns flag file for platform: $($lProdCode.id)"
                $lSoftpaqList = (Get-SoftpaqList -platform $lProdCode -os $Script:OS -osver $Script:v_OSVER -characteristic SSM) 6>&1  

                ForEach ( $iEntry in $lAddOnsList ) {
                    ForEach ( $iSoftpaq in $lSoftpaqList ) {
                        if ( [string]$iSoftpaq.Name -match $iEntry ) {
                            # CMTraceLog -Message "... found $($iSoftpaq.id) $($iSoftpaq.Name)"
                            $lSoftpaqExe = $lAddSoftpaqsFolder+'\'+$iSoftpaq.id+'.exe'
                            if ( Test-Path $lSoftpaqExe ) {
                                CMTraceLog -Message "... $($iSoftpaq.id) already downloaded - $($iSoftpaq.Name)"
                            } else {
                                $ret = (Get-Softpaq $iSoftpaq.Id) 6>&1
                                CMTraceLog -Message  "- Downloaded $($ret)"  -Type $TypeWarn
                            } # else if ( Test-Path $lSoftpaqExe ) {
                            # ALWAYS update the CVA file, even if previously downloaded
                            $ret = (Get-SoftpaqMetadataFile $iSoftpaq.Id) 6>&1
                            if ( $DebugMode ) { CMTraceLog -Message  "... Get-SoftpaqMetadataFile done: $($ret)"  -Type $TypeWarn }
                        } # if ( [string]$iSoftpaq.Name -match $iEntry )
                    } # ForEach ( $iSoftpaq in $lSoftpaqList )
                } # ForEach ( $lEntry in $lAddOnsList )

            } else {
                CMTraceLog -Message $lProdCode'Flag file found but empty'
            } # else if ( -not ([String]::IsNullOrWhiteSpace($lAddOns)) )
        } # if ( $lAddOnsFlag -and (Test-Path $lProdIDFlagFile) )
    } # else if ( $pCommonRepoFlag )
    
    Set-Location $lCurrentLoc

} # Function Download_AddOn_Softpaqs

#=====================================================================================
<#
    Function Restore_AddOn_Softpaqs
        This function will restore softpaqs from .ADDSOFTWARE to root

    expects Parameter 
        - Repository folder to sync
#>
Function Restore_AddOn_Softpaqs {
    [CmdletBinding()]
	param( $pFolder, 
            $pAddSoftpaqs )        # $True = restore Softpaqs to repo from .ADDSOTWARE folder
    
        if ( $pAddSoftpaqs ) {
            CMTraceLog -Message "... Copying named softpaqs/cva files to Repository - if selected"
            $lCmd = "robocopy.exe"
            $ldestination = "$($pFolder)"
            $lsource = "$($pFolder)\$($s_AddSoftware)"
            # Start robocopy -args "$source $destination $fileLIst $robocopyOptions" 
            Start-Process $lCmd -args "$lsource $ldestination *cva" -WindowStyle Minimized
            Start-Process $lCmd -args "$lsource $ldestination *exe" -WindowStyle Minimized
        } # if ( -not $noIniSw )

} # Restore_AddOn_Softpaqs

#=====================================================================================
<#
    Function Sync_and_Cleanup_Repository
        This function will run a Sync and a Cleanup commands from HPCMSL

    expects Parameter 
        - pGrid
        - Repository folder to sync
        - Flag : True if this is a common repo, False if an individual repo
#>
Function Sync_and_Cleanup_Repository {
    [CmdletBinding()]
	param( $pGrid, $pFolder, $pCommonFlag )

    $lCurrentLoc = Get-Location
    CMTraceLog -Message  "... [Sync_and_Cleanup_Repository] - <$pFolder> - please wait !!!" -Type $TypeNorm

    if ( Test-Path $pFolder ) {
        #--------------------------------------------------------------------------------
        # update repository softpaqs with sync command and then cleanup
        #--------------------------------------------------------------------------------
        Set-Location -Path $pFolder

        # but first let's start a look at what connections (port 443) are being used, wait 2 seconds
        #--------------------------------------------------------------------------------
        Get_HPServerIPConnection 'starting repository sync' $HPServerIPFile 2
        CMTraceLog -Message  '... calling invoke-repositorysync()' -Type $TypeNorm
        invoke-repositorysync 6>&1
        #--------------------------------------------------------------------------------
        # see if we captured a connection to HP
        Wait-Job -Name invoke-sync -ErrorAction:SilentlyContinue 
        Receive-Job -Name invoke-sync -ErrorAction:SilentlyContinue 
        Remove-Job -Name invoke-sync -ErrorAction:SilentlyContinue 

        $lOut15File = "$($HPServerIPFile)_invoke-repositorysync.TXT"
        # now, analyze for HP connection IP, if one found - should be 15.xx.xx.xx
        if ( Test-Path $lOut15File ) {
            ForEach ($line in (Get-Content $lOut15File)) { 
                if ( $DebugMode ) { CMTraceLog -Message "      $line" -Type $TypeWarn }
            }
            remove-Item $lOut15File -Force        
        } # if ( Test-Path $usingpOutFile )

        # find what sync'd from the CMSL log file for this run
        Get_ActivityLogEntries $pFolder  # get sync'd info from HPIA activity log file
        if ( Test-Path $pFolder ) {
            $lContentsHASH = (Get-FileHash -Path "$($pFolder)\.repository\contents.csv" -Algorithm MD5).Hash
        } else { 
            $lContentsHASH = $null
        }
        if ( $DebugMode ) { CMTraceLog -Message "... MD5 Hash of 'contents.csv': $($lContentsHASH)" -Type $TypeWarn }

        #--------------------------------------------------------------------------------
        CMTraceLog -Message  '... calling invoke-RepositoryCleanup()' -Type $TypeNorm
        invoke-RepositoryCleanup 6>&1
        #--------------------------------------------------------------------------------

        # see if Cleanup modified the contents.csv file 
        # - seems like (up to at least 1.6.3) RepositoryCleanup does not Modify 'contents.csv'
        #$lContentsHASH = Get_SyncdContents $pFolder       # Sync command creates 'contents.csv'
        #CMTraceLog -Message "... MD5 Hash of 'contents.csv': $($lContentsHASH)" -Type $TypeWarn

        #-----------------------------------------------------------------------------------
        # Now, we manage the 'AddOns' Softpaqs maintained in .ADDSOFTWARE
        # ... note each platform should have a flag file w/ID name if this is a common repo
        #-----------------------------------------------------------------------------------
        CMTraceLog -Message  '... calling Download_AddOn_Softpaqs()' -Type $TypeNorm
        Download_AddOn_Softpaqs $pGrid $pFolder $pCommonFlag

        CMTraceLog -Message  '... calling Restore_Softpaqs()' -Type $TypeNorm
        # next, copy all softpaqs in $s_AddSoftware subfolder to the repository 
        # ... (since it got cleared up by CMSL's "Invoke-RepositoryCleanup")
        Restore_AddOn_Softpaqs $pFolder (-not $script:noIniSw)

        CMTraceLog -Message  '... [Sync_and_Cleanup_Repository] - Done' -Type $TypeNorm

    } # if ( Test-Path $pFolder )

    Set-Location $lCurrentLoc

} # Function Sync_and_Cleanup_Repository

#=====================================================================================
<#
    Function Update_Model_Filters
        for the selected model: 
            - remove all filters
            - add filters as selected in UI 
        ******* TBD: Add Platform AddOns file to .ADDSOFTWARE
#>
Function Update_Model_Filters {
[CmdletBinding()]
	param( $pGrid,
            $pModelRepository,
            $pModelID,
            $pRow )

    $pCurrentLoc = Get-Location

    set-location $pModelRepository         # move to location of Repository to use CMSL repo commands

    if ( $Script:v_Continueon404 ) {
        Set-RepositoryConfiguration -Setting OnRemoteFileNotFound -Value LogAndContinue
    } else {
        Set-RepositoryConfiguration -Setting OnRemoteFileNotFound -Value Fail
    }
 
    #--------------------------------------------------------------------------------
    # now update filters for every category checked for the current model'
    #--------------------------------------------------------------------------------
    if ( $DebugMode ) { CMTraceLog -Message "... adding category filters" -Type $TypeDebug }

    foreach ( $cat in $v_FilterCategories ) {
        if ( $pGrid.Rows[$pRow].Cells[$cat].Value ) {
            CMTraceLog -Message  "... adding filter: -Platform $pModelID -os $OS -osver $v_OSVER -category $($cat) -characteristic ssm" -Type $TypeNoNewline
            $lRes = (Add-RepositoryFilter -platform $pModelID -os win10 -osver $v_OSVER -category $cat -characteristic ssm 6>&1)
            CMTraceLog -Message $lRes -Type $TypeWarn 
        }
    } # foreach ( $cat in $v_FilterCategories )
    #--------------------------------------------------------------------------------
    # update repository path field for this model in the grid (path is the last col)
    #--------------------------------------------------------------------------------
    $pGrid[($pGrid.ColumnCount-1),$pRow].Value = $pModelRepository
<#
    # if we are maintaining AddOns Softpaqs, see it in the checkbox (column 7, 'AddOns' cell)
    # ... first get path to Platform IID flag file
    $lAddOnsRepoFile = $pModelRepository+'\'+$s_AddSoftware+'\'+$pModelID 
    # ... next check if the user wants to sync softpaqs for this platform
    $lAddOnsFlag = $pGrid.Rows[$pRow].Cells['AddOns'].Value # column 7 has the AddOns checkmark
#>
    Set-Location -Path $pCurrentLoc

    Return $lAddOnsFlag

} # Update_Model_Filters

#=====================================================================================
<#
    Function Sync_Repos
    This function is used for command-line execution
#>
Function Sync_Repos {
[CmdletBinding()]
	param( $pGrid, $pCommonRepoFlag )

    $lCurrentSetLoc = Get-Location

    #--------------------------------------------------------------------------------
    # find out if the share exists, if not, just return
    #--------------------------------------------------------------------------------

    if ( $pCommonRepoFlag ) {
            
        # basic check to confirm the repository was configured for HPIA
        if ( !(Test-Path "$($Script:v_Root_CommonRepoFolder)\.Repository") ) {
            CMTraceLog -Message  "... Common Repository Folder selected, not initialized" -Type $TypeNorm
            return
        } 
        set-location $Script:v_Root_CommonRepoFolder                 
        Sync_and_Cleanup_Repository $pGrid $Script:v_Root_CommonRepoFolder $pCommonRepoFlag

    } else {

        # basic check to confirm the repository exists that hosts individual repos
        if ( Test-Path $Script:v_Root_IndividualRepoFolder ) {
            # let's traverse every product Repository folder
            $lProdFolders = Get-ChildItem -Path $Script:v_Root_IndividualRepoFolder | Where-Object {($_.psiscontainer)}

            foreach ( $lprodName in $lProdFolders ) {
                $lCurrentPath = "$($Script:v_Root_IndividualRepoFolder)\$($lprodName.name)"
                set-location $lCurrentPath
                Sync_and_Cleanup_Repository $pGrid $lCurrentPath $pCommonRepoFlag
            } # foreach ( $lprodName in $lProdFolders )
        } else {
            CMTraceLog -Message  "... Shared/Individual Repository Folder selected, Head repository not initialized" -Type $TypeNorm
        } # else if ( !(Test-Path $Script:v_Root_IndividualRepoFolder) ) 
    } # else if ( $Script:v_CommonRepo )

    CMTraceLog -Message  "Sync DONE" -Type $TypeSuccess
    Set-Location -Path $lCurrentSetLoc

} # Sync_Repos

#=====================================================================================
<#
    Function sync_individual_repositories
        for every selected model, go through every repository by model
            - ensure there is a valid repository
            - remove all filters
            - add filters as selected in UI
            - invoke sync and cleanup function on each folder 
#>
Function sync_individual_repositories {
[CmdletBinding()]
	param( $pGrid,
            $pRepoHeadFolder,                                            
            $pCheckedItemsList,
            $pNewModels )                                # $True = models added to list)                                      # array of rows selected
    
    CMTraceLog -Message "> sync_individual_repositories - START" -Type $TypeNorm
    if ( $DebugMode ) { CMTraceLog -Message "... for OS version: $($v_OSVER)" -Type $TypeDebug }

    #--------------------------------------------------------------------------------
    # decide if we work on single repo or individual repos
    #--------------------------------------------------------------------------------     
    init_repository $pRepoHeadFolder $false              # make sure Main repo folder exists, do NOT make it a repository

    #--------------------------------------------------------------------------------
    if ( $DebugMode ) { CMTraceLog -Message "... stepping through all selected models" -Type $TypeDebug }

    # go through every selected HP Model in the list
    for ( $i = 0; $i -lt $pGrid.RowCount; $i++ ) {
        
        $lModelId = $pGrid[1,$i].Value                      # column 1 has the Model/Prod ID
        $lModelName = $pGrid[2,$i].Value                    # column 2 has the Model name
        #$lAddOnsFlag = $pGrid.Rows[$i].Cells['AddOns'].Value # column 7 has the AddOns checkmark

        # if model entry is checked, we need to create a repository, but ONLY if $v_CommonRepo = $false
        if ( $i -in $pCheckedItemsList ) {

            CMTraceLog -Message "`n--- Updating model: $lModelName"

            $lTempRepoFolder = "$($pRepoHeadFolder)\$($lModelName)"     # this is the repo folder for this model
            init_repository $lTempRepoFolder $true
            
            $lAddOnsFlag = Update_Model_Filters $pGrid $lTempRepoFolder $lModelId $i
            #--------------------------------------------------------------------------------
            # now sync up and cleanup this individual repository
            #--------------------------------------------------------------------------------
            Sync_and_Cleanup_Repository $pGrid $lTempRepoFolder $False # sync individual repo
            #--------------------------------------------------------------------------------
            # update SCCM Repository package, if user allows
            #--------------------------------------------------------------------------------
            if ( $Script:v_UpdateCMPackages ) {
                CM_RepoUpdate $lModelName $lModelId $lTempRepoFolder
            }
            
        } # if ( $i -in $pCheckedItemsList )

    } # for ( $i = 0; $i -lt $pGrid.RowCount; $i++ )

    if ( $pNewModels ) { Update_INIModelsList $pGrid $pCommonRepoFolder $True }

    CMTraceLog -Message "< sync_individual_repositories DONE" -Type $TypeSuccess

} # Function sync_individual_repositories

#=====================================================================================
<#
    Function Sync_Common_Repository
        for every selected model, 
            - ensure there is a valid repository
            - remove all filters
            - add filters as selected in UI
            - invoke sync and cleanup function on the folder 
#>
Function Sync_Common_Repository {
[CmdletBinding()]
	param( $pGrid,                                  # the UI grid of platforms
            $pRepoFolder,                           # location of repository
            $pCheckedItemsList )                    # array of rows selected in GUI's left column
    
    $pCurrentLoc = Get-Location
    CMTraceLog -Message "> Sync_Common_Repository - START" -Type $TypeNorm
    if ( $DebugMode ) { CMTraceLog -Message "... for OS version: $($Script:v_OSVER)" -Type $TypeDebug }

    init_repository $pRepoFolder $true        # make sure Main repo folder exists, or create it, init
    CMTraceLog -Message  "... Common repository selected: $($pRepoFolder)" -Type $TypeNorm

    if ( $Script:v_KeepFilters ) {
        CMTraceLog -Message  "... Keeping existing filters in: $($pRepoFolder)" -Type $TypeNorm
    } else {
        CMTraceLog -Message  "... Removing existing filters in: $($pRepoFolder)" -Type $TypeNorm
    } # else if ( $Script:v_KeepFilters )

    if ( $DebugMode ) { CMTraceLog -Message "... stepping through selected models" -Type $TypeDebug }
    
    #-------------------------------------------------------------------------------------------
    # loop through every Model in the grid, and look for selected rows in UI (left cell checked)
    #-------------------------------------------------------------------------------------------
    for ( $item = 0; $item -lt $pGrid.RowCount; $item++ ) {
        
        $lModelSelected = $pGrid[0,$item].Value
        $lModelId = $pGrid[1,$item].Value                # column 1 has the Model/Prod ID
        $lModelName = $pGrid[2,$item].Value              # column 2 has the Model name
        #$lAddOnsFlag = $pGrid.Rows[$item].Cells['AddOns'].Value # column 7 has the AddOns checkmark

        #---------------------------------------------------------------------------------
        # Remove existing filter for this platform, unless the KeepFilters checkbox is set
        #---------------------------------------------------------------------------------
        if ( -not $Script:v_KeepFilters ) {
            $lres = (Remove-RepositoryFilter -platform $lModelID -yes 6>&1)      
            if ( $debugMode ) { CMTraceLog -Message "... removed filters for: $($lModelID)" -Type $TypeWarn }
        } # if ( $Script:v_KeepFilters )

        #if ( $item -in $pCheckedItemsList ) {
        if ( $lModelSelected ) {

            CMTraceLog -Message "--- Updating model: $lModelName"
            #--------------------------------------------------------------------------------
            # update repo filters from grid
            #--------------------------------------------------------------------------------
            $lAddOnsFlag = Update_Model_Filters $pGrid $pRepoFolder $lModelId $item
            #--------------------------------------------------------------------------------
            # update SCCM Repository package, if enabled in the UI checkmark
            #--------------------------------------------------------------------------------
            if ( $Script:v_UpdateCMPackages ) { CM_RepoUpdate $lModelName $lModelId $pRepoFolder }
        } # if ( $item -in $pCheckedItemsList )
    } # for ( $item = 0; $item -lt $pGrid.RowCount; $item++ )

    #-----------------------------------------------------------------------------------
    # we are done checking every model for filters, so now do a Softpaq Sync and cleanup
    #-----------------------------------------------------------------------------------
    Sync_and_Cleanup_Repository $pGrid $pRepoFolder $True # $True = sync a common repository

    #if ( $pNewModels ) { Update_INIModelsList $pGrid $pRepoFolder $False }   # $False = head of individual repos

    #--------------------------------------------------------------------------------
    Set-Location -Path $pCurrentLoc

    CMTraceLog -Message "< Sync_Common_Repository DONE" -Type $TypeSuccess

} # Function Sync_Common_Repository

#=====================================================================================
<#
    Function clear_grid
        clear all checkmarks and last path column '8', except for SysID and Model columns

#>
Function clear_grid {
    [CmdletBinding()]
	param( $pDataGrid )                             

    if ( $DebugMode ) { CMTraceLog -Message '> clear_grid' -Type $TypeNorm }

    # scan every row in the grid
    for ( $row = 0 ; $row -lt $pDataGrid.RowCount ; $row++ ) {
        # for each row, clear each cell
        for ( $col = 0 ; $col -lt $pDataGrid.ColumnCount ; $col++ ) {
            # clear all columns, except 1,2 - 'SysID' and 'System Name'
            if ( $col -in @(0,3,4,5,6,7) ) {
                $pDataGrid[$col,$row].value = $False                       # clear checkmarks
            } else {
                # clear the 'repository path' cell (last column)
                if ( $pDataGrid.columns[$col].Name -match 'Repository' ) {
                    $pDataGrid[$col,$row].value = ''                       # clear path text field
                }
            } # else if ( $col -in @(0,3,4,5,6,7) )
        } # for ( $col = 0 ; $col -lt $pDataGrid.ColumnCount ; $col++ )
    } # for ( $row = 0 ; $row -lt $pDataGrid.RowCount ; $row++ ) 
    
    if ( $DebugMode ) { CMTraceLog -Message '< clear_grid' -Type $TypeNorm }

} # Function clear_grid

#=====================================================================================
<#
    Function list_filters
        List filters to log file... function ONLY called from runstring option
        ... lists filters for both single, common or individual repositories
#>
Function list_filters {

    $lCurrentSetLoc = Get-Location

    if ( $Script:v_CommonRepo ) {
        # basic check to confirm the repository was configured for HPIA
        if ( !(Test-Path "$($Script:v_Root_CommonRepoFolder)\.Repository") ) {
            CMTraceLog -Message "... Common Repository Folder selected, not initialized" -Type $TypeNorm 
            "... Common Repository Folder selected, not initialized for HPIA"
            return
        } 
        set-location $Script:v_Root_CommonRepoFolder
        $lProdFilters = (get-repositoryinfo).Filters
        foreach ( $filterSetting in $lProdFilters ) {
            $lMsg = "`t-Platform $($filterSetting.platform) -OS $($filterSetting.operatingSystem) -Category $($filterSetting.category) -Characteristic $($filterSetting.characteristic)"
            if ( $Products ) {
                if ( $filterSetting.platform -in $Products ) {
                    CMTraceLog -Message $lMsg -Type $TypeNorm
                } # if ( $filterSetting.platform -in $Products )
            } else {
                CMTraceLog -Message $lMsg -Type $TypeNorm
            }
        } # foreach ( $filterSetting in $lProdFilters )
    } else {
        # basic check to confirm the head repository exists
        if ( !(Test-Path $Script:v_Root_IndividualRepoFolder) ) {
            CMTraceLog -Message "... Shared/Individual Repository Folder selected, requested folder NOT found" -Type $TypeNorm
            "... Shared/Individual Repository Folder selected, Head repository not initialized"
            return
        } 
        set-location $Script:v_Root_IndividualRepoFolder | Where-Object {($_.psiscontainer)}

        # let's traverse every product Repository folder
        $lProdFolders = Get-ChildItem -Path $Script:v_Root_IndividualRepoFolder

        foreach ( $lprodName in $lProdFolders ) {
            set-location "$($Script:v_Root_IndividualRepoFolder)\$($lprodName.name)"
            $lProdFilters = (get-repositoryinfo).Filters
            $lprods = $lProdFilters.platform.split()       # if multiple filters, each will have same ProdCode, so just show one [0]
            $lcharacteristics = $lProdFilters.characteristic | Get-Unique
            $lMsg = "`t-Platform $($lprods[0]) -OS $($lProdFilters.operatingSystem) -Category $($lProdFilters.category) -Characteristic $($lcharacteristics)"
            if ( $Products ) {
                if ( $lprods[0] -in $Products ) {
                    CMTraceLog -Message $lMsg -Type $TypeNorm
                } # if ( $lprods[0] -in $Products )
            } else {
                CMTraceLog -Message $lMsg -Type $TypeNorm
            }
        } # foreach ( $lprodName in $lProdFolders )
    } # else if ( $Script:v_CommonRepo )

    Set-Location -Path $lCurrentSetLoc

} # list_filters

#=====================================================================================
<#
    Function Get_CommonRepofilters
        Retrieves category filters from a single, common repository
        ... and populates the Grid appropriately
        ... including if AddOns are being maintained
#>
Function Get_CommonRepofilters {
    [CmdletBinding()]
	param( $pGrid,                              # array of row lines that are checked
        $pCommonFolder,
        $pRefreshGrid )                             # if $true, refresh grid,  $false, just list filters

    if ( $DebugMode ) { CMTraceLog -Message '> Get_CommonRepofilters' -Type $TypeNorm }

    # see if the repository was configured for HPIA
    if ( !(Test-Path "$($pCommonFolder)\.Repository") ) {
        CMTraceLog -Message "... Repository Folder not initialized" -Type $TypeWarn
        if ( $DebugMode ) { CMTraceLog -Message '< Get_CommonRepofilters - Done' -Type $TypeWarn }
        return
    } 

    if ( $pRefreshGrid ) {
        CMTraceLog -Message "... Refreshing Grid from Common Repository: ''$pCommonFolder''" -Type $TypeNorm
        clear_grid $pGrid
    } else {
        CMTraceLog -Message "... Filters from Common Repository ...$($pCommonFolder)" -Type $TypeNorm
    }
    #---------------------------------------------------------------
    # Let's now get all the HPIA filters from the repository
    #---------------------------------------------------------------
    set-location $pCommonFolder
    $lProdFilters = (get-repositoryinfo).Filters   # get the list of filters from the repository

    ### example filter returned by 'get-repositoryinfo' CMSL command: 
    ###
    ### platform        : 8438
    ### operatingSystem : win10:2004 win10:2004
    ### category        : BIOS firmware
    ### releaseType     : *
    ### characteristic  : ssm

    foreach ( $filter in $lProdFilters ) {
        # check each row SysID against the Filter Platform ID
        for ( $i = 0; $i -lt $pGrid.RowCount; $i++ ) {
            $lPlatform = $filter.platform

            if ( $pRefreshGrid ) {
                if ( $lPlatform -eq $pGrid[1,$i].value ) {
                    # we matched the row/SysId with the Filter

                    # ... so let's add each category in the filter to the Model in the GUI
                    foreach ( $cat in  ($filter.category.split(' ')) ) {
                        $pGrid.Rows[$i].Cells[$cat].Value = $true
                    }
                    $pGrid[0,$i].Value = $true   # check the selection column 0
                    $pGrid[($pGrid.ColumnCount-1),$i].Value = $pCommonFolder # ... and the Repository Path
                
                    # if we are maintaining AddOns Softpaqs, show it in the checkbox
                    $lAddOnsRepoFile = $pCommonFolder+'\'+$s_AddSoftware+'\'+$lPlatform 
                    if (Test-Path $lAddOnsRepoFile) {
                        if ( -not ($pGrid.Rows[$i].Cells['AddOns'].Value) ) { # expose only once/platform
                            $pGrid.Rows[$i].Cells['AddOns'].Value = $True
                            [array]$lAddOns = Get-Content $lAddOnsRepoFile
                            CMTraceLog -Message "...$lPlatform - Additional Softpaqs: $lAddOns" -Type $TypeWarn  
                        } #                    
                    } # if ( Test-Path $lAddOnRepoFile )
                } # if ( $lPlatform -eq $pGrid[1,$i].value )
            } else {
                List_Filters
            } # else if ( $pRefreshGrid )
        } # for ( $i = 0; $i -lt $pGrid.RowCount; $i++ )
    } # foreach ( $filter in $lProdFilters )

    [void]$pGrid.Refresh
    CMTraceLog -Message "... Refreshing Grid ...DONE" -Type $TypeSuccess

    if ( $DebugMode ) { CMTraceLog -Message '< Get_CommonRepofilters()' -Type $TypeNorm }
    
} # Function Get_CommonRepofilters

#=====================================================================================
<#
    Function Get_IndividualRepofilters
        Retrieves category filters from the share for each selected model
        ... and populates the Grid appropriately
        Parameters:
            $pGrid                          The models grid in the GUI
            $pRepoLocation                      Where to start looking for repositories
            $pRefreshGrid                       $True to refresh the GUI grid, $false to just list filters
#>
Function Get_IndividualRepofilters {
    [CmdletBinding()]
	param( $pGrid,                                  # array of row lines that are checked
        $pRepoLocation,
        $pRefreshGrid )                                 # if $false, just list filters, don't refresh grid

    if ( $DebugMode ) { CMTraceLog -Message '> Get_IndividualRepofilters' -Type $TypeNorm }

    set-location $pRepoLocation
    if ( $pRefreshGrid ) {
        CMTraceLog -Message '... Refreshing Grid from Individual Repositories ...' -Type $TypeNorm
    }
    #--------------------------------------------------------------------------------
    # now check for each product's repository folder
    # if the repo is created, then check the category filters
    #--------------------------------------------------------------------------------
    for ( $i = 0; $i -lt $pGrid.RowCount; $i++ ) {

        $lModelId = $pGrid[1,$i].Value                     # column 1 has the Model/Prod ID
        $lModelName = $pGrid[2,$i].Value                   # column 2 has the Model name
        $lTempRepoFolder = "$($pRepoLocation)\$($lModelName)"  # this is the repo folder for this model
        
        if ( Test-Path $lTempRepoFolder ) {
            set-location $lTempRepoFolder                      # move to location of Repository to use CMSL repo commands
                                    
            $lProdFilters = (get-repositoryinfo).Filters

            foreach ( $platform in $lProdFilters ) {
                if ( $pRefreshGrid ) {                    
                    if ( $DebugMode ) { CMTraceLog -Message "... populating filter categories ''$($lModelName)''" -Type $TypeNorm }
                    foreach ( $cat in  ($lProdFilters.category.split(' ')) ) {
                        $pGrid.Rows[$i].Cells[$cat].Value = $true
                    }
                    #--------------------------------------------------------------------------------
                    # show repository path for this model (model is checked) (col 8 is links col)
                    #--------------------------------------------------------------------------------
                    $pGrid[0,$i].Value = $true
                    $pGrid[($dataGridView.ColumnCount-1),$i].Value = $lTempRepoFolder

                    # if we are maintaining AddOns Softpaqs, show it in the checkbox
                    $lAddOnRepoFile = $lTempRepoFolder+'\'+$s_AddSoftware+'\'+$lProdFilters.platform   
                    if ( Test-Path $lAddOnRepoFile ) { 
                        $pGrid.Rows[$i].Cells['AddOns'].Value = $True
                        $lPlatform = $($lProdFilters.platform) | Get-Unique
                        $lMsg = "...Additional Softpaqs Enabled for Platform ''$lPlatform'': "+$Script:HPModelsTable[$i].AddOns
                        CMTraceLog -Message $lMsg -Type $TypeWarn 
                    }
                } else {                           # just list the filters, no grid refresh
                    CMTraceLog -Message "... Platform $($lProdFilters.platform) ... $($lProdFilters.operatingSystem) $($lProdFilters.category) $($lProdFilters.characteristic) - @$($lTempRepoFolder)" -Type $TypeWarn
                }
            } # foreach ( $platform in $lProdFilters )

        } # if ( Test-Path $lTempRepoFolder )
    } # for ( $i = 0; $i -lt $pModelsList.RowCount; $i++ ) 

    if ( $DebugMode ) { CMTraceLog -Message '< Get_IndividualRepofilters]' -Type $TypeNorm }
    
} # Function Get_IndividualRepofilters

#=====================================================================================
<#
    Function Empty_Grid
        removes (empties out) all Model entries (rows) from the grid
#>
Function Empty_Grid {
    [CmdletBinding()]
	param( $pGrid )

    for ( $i = $pGrid.RowCount; $i -gt 0; $i-- ) {
        $pGrid.Rows.RemoveAt($i-1)
    } # for ( $i = 0; $i -lt $pGridList.RowCount; $i++ )            

} # Function Empty_Grid

#=====================================================================================
<#
    Function Browse_IndividualRepos
        Browse to find existing or create new repository
        Paramaters: $pGrid = pointer to model grid in UI
                    $pCurrentRepository
#>
Function Browse_IndividualRepos {
    [CmdletBinding()]
	param( $pGrid, $pCurrentRepository )  

    # let's find the repository to import
    $lbrowse = New-Object System.Windows.Forms.FolderBrowserDialog
    $lbrowse.SelectedPath = $pCurrentRepository         
    $lbrowse.Description = "Select an HPIA Head Repository Folder"
    $lbrowse.ShowNewFolderButton = $true
                                  
    if ( $lbrowse.ShowDialog() -eq "OK" ) {
        $lRepository = $lbrowse.SelectedPath
        CMTraceLog -Message  "... clearing Model list" -Type $TypeNorm
        Empty_Grid $pGrid
        CMTraceLog -Message  "... moving to location: $($lRepository)" -Type $TypeNorm
        Set-Location $lRepository
        $lDirectories = (Get-ChildItem -Directory)     
        # Let's search for repositories (subfolders) at this location
        # add a row in the DataGrid for every repository (e.g. model) we find
        foreach ( $lFolder in $lDirectories ) {
            $lRepoFolder = $lRepository+'\'+$lFolder
            Set-Location $lRepoFolder
            Try {
                $lRepoFilters = (Get-RepositoryInfo).Filters
                CMTraceLog -Message  "... Repository Found ($($lRepoFolder)) -- adding to grid" -Type $TypeNorm
                # obtain platform SysID from filters in the repository
                [array]$lRepoPlatform = $lRepoFilters.platform
                [void]$pGrid.Rows.Add(@( $true, $lRepoPlatform, $lFolder ))
            } Catch {
                CMTraceLog -Message  "... $($lRepoFolder) is NOT a Repository" -Type $TypeNorm
            } # Catch
        } # foreach ( $lFolder in $lDirectories )
        
        if ( $DebugMode ) { CMTraceLog -Message  "... Retrieving Filters from repositories" }
        Get_IndividualRepofilters $pGrid $lRepository $True
        if ( $DebugMode ) { CMTraceLog -Message  "... calling Update_UIandINI()" }
        Update_UIandINI $lRepository $False              # $False = using individual repositories
        if ( $DebugMode ) { CMTraceLog -Message  "... calling Update_INIModelsList()" }
        Update_INIModelsList $pGrid $lRepository $False  # $False = treat as head of individual repositories
        CMTraceLog -Message "Browse Individual Repositories Done ($($lRepository))" -Type $TypeSuccess
    } else {
        $lRepository = $pCurrentRepository
    } # else if ( $lbrowse.ShowDialog() -eq "OK" )
    
    $lbrowse.Dispose()

} # Function Browse_IndividualRepos

#=====================================================================================
<#
    Function Import_IndividualRepos
        - Populate list of Platforms from Individual Repositories starting at the 1st argument
        - also, update INI file about imported repository models if user agrees
        Paramaters: $pGrid = pointer to model grid in UI
                    $pCurrentRepository = head of individual repository folders
                    $pINIModelsList = list of models from INI.ps1 file
                    
        returns: object w/Folder picked for repository, or $null if user cancelled
                Folder return value contains True/Fail entry [0] + folder name [1]
#>
Function Import_IndividualRepos {
    [CmdletBinding()]
	param( $pGrid, $pCurrentRepository, $pINIModelsList )                                

    CMTraceLog -Message "> Import_IndividualRepos($($pCurrentRepository))" -Type $TypeNorm

    if ( -not (Test-Path $pCurrentRepository) ) {
        CMTraceLog "Individual Root Path from INI file not found: $($pCurrentRepository) - will create"  -Type $TypeWarn
        Init_Repository $pCurrentRepository $False            # False = create root folder only
    }
    CMTraceLog -Message  "... clearing Model list" -Type $TypeNorm
    Empty_Grid $pGrid
    CMTraceLog -Message  "... moving to location: $($pCurrentRepository)" -Type $TypeNorm
    Set-Location $pCurrentRepository

    $lDirectories = (Get-ChildItem -Directory)                    # each subfolder is an Individual model repository

    if ( $lDirectories.count -eq 0 ) {
        #--------------------------------------------------------------------------------
        # if the repository has no subfolders, use the models from the INI file list
        #--------------------------------------------------------------------------------        
        Populate_Grid_from_INI $pGrid $pCurrentRepository $pINIModelsList
    } else {
        #--------------------------------------------------------------------------------
        # else populate from each valid individual repository subfolder
        #--------------------------------------------------------------------------------
        foreach ( $lProdName in $lDirectories ) {
            $lRepoFolderFullPath = $pCurrentRepository+'\'+$lProdName
            Set-Location $lRepoFolderFullPath
            Try {
                CMTraceLog -Message  "... Checking repository: $($lRepoFolderFullPath) -- adding to grid" -Type $TypeNorm
                $lRepoFilters = (Get-RepositoryInfo).Filters | Get-Unique
                $lRow = $pGrid.Rows.Add(@( $true, $lRepoFilters.platform, $lProdName ))
                $lAddOnsFlagFile = $lRepoFolderFullPath+'\.ADDSOFTWARE\'+$lRepoFilters.platform
                if ( Test-Path -Path $lAddOnsFlagFile ) {
                    if ( -not ( [String]::IsNullOrWhiteSpace((Get-content $lAddOnsFlagFile)) ) ) {
                        $pGrid.rows[$lRow].Cells['AddOns'].Value = $True
                    }
                } # if ( Test-Path -Path $lAddOnsFlagFile )
            } Catch {
                CMTraceLog -Message  "... NOT a Repository: ($($lRepoFolderFullPath))" -Type $TypeNorm
            } # Catch
        } # foreach ( $lProdName in $lDirectories )
        CMTraceLog -Message  "... Getting Filters from repositories"
        Get_IndividualRepofilters $pGrid $pCurrentRepository $True
        CMTraceLog -Message  "... Updating UI and INI Path"
        Update_UIandINI $pCurrentRepository $False    # $False = individual repos
        Update_INIModelsList $pGrid $pCurrentRepository $False  # $False = this is for Individual repositories
        CMTraceLog -Message "< Import_IndividualRepos() Done" -Type $TypeSuccess
    } # else if ( $lDirectories.count -eq 0 )

    if ( $DebugMode ) { CMTraceLog -Message '< Import_IndividualRepos]' -Type $TypeNorm }

} # Function Import_IndividualRepos

#=====================================================================================
<#
    Function Browse_CommonRepo
        Browse to find existing or create new repository
        Paramaters: $pGrid = pointer to model grid in UI
                    $pCurrentRepository
        returns: object w/Folder picked for repository, or $null if user cancelled
                Folder return value contains True/Fail entry [0] + folder name [1]
#>
Function Browse_CommonRepo {
    [CmdletBinding()]
	param( $pGrid, $pCurrentRepository )                                

    # let's find the repository to import or use the current

    $lbrowse = New-Object System.Windows.Forms.FolderBrowserDialog
    $lbrowse.SelectedPath = $pCurrentRepository        # start with the repo listed in the INI file
    $lbrowse.Description = "Select an HPIA Common/Shared Repository"
    $lbrowse.ShowNewFolderButton = $true

    if ( $lbrowse.ShowDialog() -eq "OK" ) {
        $lRepository = $lbrowse.SelectedPath

        Empty_Grid $pGrid
        #--------------------------------------------------------------------------------
        # find out if the share exists and has the context for HPIA, if not, just return
        #--------------------------------------------------------------------------------
        Try {
            Set-Location $lRepository
            $lProdFilters = (Get-RepositoryInfo).Filters
             # ... populate grid with platform SysIDs found in the repository
             # ... if this fails then this is not a proper repository, so go to 'Catch' to initialize
            [array]$lRepoPlatforms = $lProdFilters.platform | Get-Unique
            for ( $i = 0 ; $i -lt $lRepoPlatforms.count ; $i++) {
                [array]$lProdName = Get-HPDeviceDetails -Platform $lRepoPlatforms[$i]
                #$row = @( $false, $lRepoPlatforms[$i], $lProdName[0].Name )  
                [void]$pGrid.Rows.Add(@( $false, $lRepoPlatforms[$i], $lProdName[0].Name ))
            } # for ( $i = 0 ; $i -lt $lRepoPlatforms.count ; $i++)
            Get_CommonRepofilters $pGrid $lRepository $True
            #Update_UIandINI $lRepository $True   # $True = common repo

            Update_INIModelsList $pGrid $lRepository $True  # $True = this is a common repository
            CMTraceLog -Message "... found HPIA repository $($lRepository)" -Type $TypeNorm
        } Catch {
            # lets initialize repository but maintain existing models in the grid
            # ... in case we want to use some of them in the new repo
            init_repository $lRepository $true
            clear_grid $pGrid
            CMTraceLog -Message "... $($lRepository) initialized " -Type $TypeNorm
        } # catch
        Update_UIandINI $lRepository $True   # $True = common repo
        CMTraceLog -Message "Browse Common Repository Done ($($lRepository))" -Type $TypeSuccess
    } else {
        $lRepository = $pCurrentRepository
    } # else if ( $lbrowse.ShowDialog() -eq "OK" )

    #return [string]$lRepository

} # Function Browse_CommonRepo

#=====================================================================================
<#
    Function Import_CommonRepo
        If this is an existing repository with filters, show the contents in the grid
            and update INI file about imported repository

        else, populate grid from INI's platform list

        Paramaters: $pGrid = pointer to model grid in UI
                    $pCurrentRepository
                    $pModelsTable               --- current models table to use if nothing in repository
#>
Function Import_CommonRepo {
    [CmdletBinding()]
	param( $pGrid, $pCurrentRepository, $pModelsTable )                                

    CMTraceLog -Message "> Import_CommonRepo($($pCurrentRepository))" -Type $TypeNorm

    if ( ([string]::IsNullOrEmpty($pCurrentRepository)) -or (-not (Test-Path $pCurrentRepository)) ) {
        CMTraceLog -Message  "... No Repository to import" -Type $TypeWarn
        return
    }
    #--------------------------------------------------------------------------------
    # at this point, we know the respository folder exists, so
    # ... find out if it is new (e.g., no filters yet): populate from INI models
    #--------------------------------------------------------------------------------
    Set-Location $pCurrentRepository
    Try {
        # Get the CMSL filters from the repository
        $lProdFilters = (Get-RepositoryInfo).Filters
        CMTraceLog -Message  "... valid HPIA Repository Found" -Type $TypeNorm
        CMTraceLog -Message  "... clearing Model list (calling Emptry_Grid())" -Type $TypeNorm
        Empty_Grid $pGrid
        CMTraceLog -Message  "... finding filters in Repository" -Type $TypeNorm

        if ( $lProdFilters.count -eq 0 ) {
            CMTraceLog -Message  "... no filters found, so populate from INI file (calling Populate_Grid_from_INI())" -Type $TypeNorm
            Populate_Grid_from_INI $pGrid $pCurrentRepository $pModelsTable
        } else {
            #--------------------------------------------------------------------------------
            # let's update the grid from in the repository filters 
            # ... first, get the (unique) list of platform SysIDs in the repository
            #--------------------------------------------------------------------------------
            [array]$lRepoPlatforms = $lProdFilters.platform | Get-Unique
            # next, add each product to the grid, to then populate with the filters
            CMTraceLog -Message  "... Adding platforms to Grid from repository"
            for ( $i = 0 ; $i -lt $lRepoPlatforms.count ; $i++) {
                # ... treat as array, as some systems have common platform IDs, so we'll pick the 1st entry
                [array]$lProdName = Get-HPDeviceDetails -Platform $lRepoPlatforms[$i]
                [void]$pGrid.Rows.Add(@( $false, $lRepoPlatforms[$i], $lProdName[0].Name ))
            } # for ( $i = 0 ; $i -lt $lRepoPlatforms.count ; $i++)
             
            CMTraceLog -Message  "... Obtaining Filters from repository" -Type $TypeNorm
            Get_CommonRepofilters $pGrid $pCurrentRepository $True
            CMTraceLog -Message  "... Updating UI and INI file from filters" -Type $TypeNorm
            Update_UIandINI $pCurrentRepository $True     # $True = common repo
            CMTraceLog -Message  "... Updating models in INI file" -Type $TypeNorm
            Update_INIModelsList $pGrid $pCurrentRepository $True  # $False means treat as head of individual repositories
        } # if ( $lProdFilters.count -gt 0 )
        CMTraceLog -Message "< Import Repository() Done" -Type $TypeSuccess
    } Catch {
        CMTraceLog -Message  "< Repository Folder ($($pCurrentRepository)) not initialized for HPIA" -Type $TypeWarn
    } # Catch
   

} # Function Import_CommonRepo

#=====================================================================================
<#
    Function Check_PlatformsOSVersion
    here we check the OS version, if supported by any platform checked in list
#>
Function Check_PlatformsOSVersion  {
    [CmdletBinding()]
	param( $pGrid,
            $pOSVersion )

    if ( $DebugMode ) { CMTraceLog -Message '> Check_PlatformsOSVersion]' -Type $TypeNorm }
    
    CMTraceLog -Message "Checking support for Win10/$($pOSVersion) for selected platforms" -Type $TypeNorm

    # search thru the table entries for checked items, and see if each product
    # has support for the selected OS version

    for ( $i = 0; $i -lt $pGrid.RowCount; $i++ ) {

        if ( $pGrid[0,$i].Value ) {
            $lPlatform = $pGrid[1,$i].Value
            $lPlatformName = $pGrid[2,$i].Value

            # get list of OS versions supported by this platform and OS version
            $lOSList = get-hpdevicedetails -platform $lPlatform -OSList -ErrorAction Continue
            
            if ( $pOSVersion -in ($lOSList).OperatingSystemRelease ) {
                CMTraceLog -Message "-- OS supported: $($lPlatform)/$($lPlatformName)" -Type $TypeNorm
            } else {
                # MS FEATURE: handle discrepancies between new OS Version Naming conventions
                switch ( $pOSVersion ) {
                    '2009' { $pOSVersion = '20H2' }
                    '2104' { $pOSVersion = '21H1' }
                    '2109' { $pOSVersion = '21H2' }
                } # switch
                if ( $pOSVersion -in ($lOSList).OperatingSystemRelease ) {
                    CMTraceLog -Message "-- OS version supported: $($lPlatform)/$($lPlatformName)" -Type $TypeNorm
                    $Script:v_OSVER = $pOSVersion    # reset to the other version
                } else {
                    CMTraceLog -Message "-- OS version NOT supported by: $($lPlatform)/$($lPlatformName)" -Type $TypeError
                }
            } # else if ( $pOSVersion -in ($lOSList).OperatingSystemRelease )
        } # if ( $dataGridView[0,$i].Value )  

    } # for ( $i = 0; $i -lt $pGrid.RowCount; $i++ )
    
    if ( $DebugMode ) { CMTraceLog -Message '< Check_PlatformsOSVersion]' -Type $TypeNorm }

} # Function Check_PlatformsOSVersion

#=====================================================================================
<#
    Function Mod_INISetting
    Function modifies a single line in the INI.ps1 file
    Searches for a setting name (find), change it to a new setting (replace)
    Parameters:
        $pFile - name of the ini file to modify
        $pfind - string to search for in the file, typically a PS variable name
        $pReplace - the line to replace the found line with
        $pText - the comment for output to log file
#>
Function Mod_INISetting {
    [CmdletBinding()]
	param( $pFile, $pfind, $pReplace, $pText )

    (Get-Content $pFile) | Foreach-Object {if ($_ -match $pfind) {$preplace} else {$_}} | Set-Content $pFile
    CMTraceLog -Message $pText -Type $TypeNorm
    
} # Mod_INISetting

#=====================================================================================
<#
    Function Update_INIModelsList
        Create text file with $HPModelsTable array from current checked grid entries
        ... and Addons from repository's model flag file(s)
        Parameters
            $pGrid
            $pRepositoryFolder
            $pCommonFlag
#>
Function Update_INIModelsList {
    [CmdletBinding()]
	param( $pGrid,
        $pRepositoryFolder,            # required to find the platform ID flag files hosted in .ADDSOFTWARE
        $pCommonFlag )                 # True = Common/shared repository

    if ( $DebugMode ) { CMTraceLog -Message '> Update_INIModelsList' -Type $TypeNorm }

    $lModelsListINIFile = $Script:IniFIleFullPath
    # -------------------------------------------------------------------
    # create list of models from the grid - assumes GUI grid is populated
    # -------------------------------------------------------------------
    $lModelsList = @()
    for ( $iRow = 0; $iRow -lt $pGrid.RowCount; $iRow++ ) {
        $lModelId = $pGrid[1,$iRow].Value             # column 1 has the Platform ID
        $lModelName = $pGrid[2,$iRow].Value           # column 2 has the Model name
        # ---------------------------------------------------------------------------------------
        # create an array string of softpaq names to add to the platform entry in the models list
        # ... note that the Flag ID file could exist and be empty
        # ---------------------------------------------------------------------------------------
        $lAddModel = "`n`t@{ ProdCode = '$($lModelId)'; Model = '$($lModelName)' }"
        $lModelsList += $lAddModel                      # add the entry to the model list
    } # for ( $iRow = 0; $iRow -lt $pGrid.RowCount; $iRow++ )
    CMTraceLog -Message '... Created HP Models List' -Type $TypeNorm
    # ---------------------------------------------------------
    # Now, replace HPModelTable in INI file with list from grid
    # ---------------------------------------------------------
    if ( Test-Path $lModelsListINIFile ) {
        
        $lListHeader = '\$HPModelsTable = .*'
        $lProdsEntry = '.*@{ ProdCode = .*'
        # remove existing model entries from file
        Set-Content -Path $lModelsListINIFile -Value (get-content -Path $lModelsListINIFile | 
            Select-String -Pattern $lProdsEntry -NotMatch)
        # add new model lines
        (get-content $lModelsListINIFile) -replace $lListHeader, "$&$($lModelsList)" | 
            Set-Content $lModelsListINIFile
        CMTraceLog -Message "... Updated Models File $lModelsListINIFile" -Type $TypeNorm
    } else {
        CMTraceLog -Message " ... INI file not updated - didn't find" -Type $TypeWarn
    } # else if ( Test-Path $lRepoLogFile )

    if ( $DebugMode ) { CMTraceLog -Message '< Update_INIModelsList' -Type $TypeNorm }
    
} # Function Update_INIModelsList 

#=====================================================================================
<#
    Function Populate_Grid_from_INI
    This is the MAIN function with a Gui that sets things up for the user
#>
Function Populate_Grid_from_INI {
    [CmdletBinding()]
	param( $pGrid,
            $pCurrentRepository,
            $pModelsTable )
<#
    populate with all the HP Models listed in the ini file
    excample line: 
    @{ ProdCode = '8723'; Model = 'HP ZBook Firefly 14 G7 Mobile Workstation' }
#>
    CMTraceLog -Message "... Populating Grid from INI file's HPModels list" -Type $TypeNorm
    $pModelsTable | 
        ForEach-Object {
            # populate checkmark, ProdId, Model Name       
            [void]$pGrid.Rows.Add( @( $False, $_.ProdCode, $_.Model) )
            # also, add AddOns Platform flag file w/contents, if setting exists for the model
            $lFlagFile = $_.Repository+'\'+$_.ProdCode
            Write-Host $lFlagFile
        } # ForEach-Object

        $pGrid.Refresh()

} # Populate_Grid_from_INI

#=====================================================================================
<#
    Function Update_UIandINI
    Update UI elements (selected path, default selection, and INI path settings)
#>
Function Update_UIandINI {
    [CmdletBinding()]
	param( $pNewPath,                   # set the appropriate global folder name to this
        $pCommonFlag )                  # set $True if $v_CommonRepo = $True      
    # ---------------------------------------------------------
    # Update INI.pse file w/new settings
    # ---------------------------------------------------------
    if ( $pCommonFlag ) {
        $CommonRadioButton.Checked = $True
        $CommonPathTextField.Text = $pNewPath ; 
        $CommonPathTextField.BackColor = $BackgroundColor
        $SharePathTextField.BackColor = ""
        $Script:v_Root_CommonRepoFolder = $pNewPath
        $find = "^[\$]v_Root_CommonRepoFolder"
        $replace = "`$v_Root_CommonRepoFolder = ""$pNewPath"""
    } else {
        $SharedRadioButton.Checked = $True
        $SharePathTextField.Text = $pNewPath
        $SharePathTextField.BackColor = $BackgroundColor
        $CommonPathTextField.BackColor = ""
        $Script:v_Root_IndividualRepoFolder = $pNewPath
        $find = "^[\$]v_Root_IndividualRepoFolder"
        $replace = "`$v_Root_IndividualRepoFolder = ""$pNewPath"""
    } # else if ( $pCommonFlag )
    # ---------------------------------------------------------
    # Update INI.ps1 file w/new repository setting
    # ---------------------------------------------------------
    Mod_INISetting $IniFIleFullPath $find $replace "... Updating '$($Script:IniFile)' setting: ''$($replace)''"
    # ---------------------------------------------------------
    # Update INI.ps1 file Comnon Repo flag setting
    # ---------------------------------------------------------
    $Script:v_CommonRepo = $pCommonFlag
    $find = "^[\$]v_CommonRepo"
    $replace = "`$v_CommonRepo = `$$pCommonFlag"  # set up the replacing string to either $false or $true from ini file
    Mod_INISetting $IniFIleFullPath $find $replace "... Updating '$($Script:IniFile)' setting: ''$($replace)''"

} # Function Update_UIandINI

#=====================================================================================
<#
    Function Manage_IDFlag_File
        Creates and updates the Pltaform ID flag file in the repository's .ADDSOFTWARE folder
        If Flag file exists and user 'unselects' AddOns, we move the file to a backup file 'ABCD_bk'
    Parameters                 
        $pPath
        $pSysID
        $pModelName
        [array]$pAddOns
        $pCreateFile -- flag                # $True to create the flag file, $False to remove (renamed)
#>
Function Manage_IDFlag_File {
    [CmdletBinding()]
	param( $pPath,
            $pSysID,                        # 4-digit hex-code platform/motherboard ID
            $pModelName,
            [array]$pAddOns,                # $Null means empty flag file (nothing to add)
            $pCreateFile)                   # $True = create file, $False = remove/rename file
    if ( $Script:v_CommonRepo ) {
        $lFlagFile = $pPath+'\'+$s_AddSoftware+'\'+$pSysID
    } else {
        $lFlagFile = $pPath+'\'+$pModelName+'\'+$s_AddSoftware+'\'+$pSysID
    }
    if ( $pCreateFile ) {
        $lMsg = "... $pSysID - Enabling download of AddOns Softpaqs "
        if ( -not (Test-Path $lFlagFile ) ) { 
            if ( Test-Path $lFlagFile'_BK' ) { 
                Move-Item $lFlagFile'_BK' $lFlagFile -Force 
                [array]$lAddOnsList = Get-Content $lFlagFile
                if ( -not ([String]::IsNullOrWhiteSpace($lAddOnsList)) ) { $lMsg += ": ... $lAddOnsList" }
            } else {
                New-Item $lFlagFile 
                if ( $pAddOns.count -gt 0 ) { 
                    $pAddOns[0] | Out-File -FilePath $lFlagFile 
                    For ($i=1; $i -le $pAddOns.count; $i++) { $pAddOns[$i] | Out-File -FilePath $lFlagFile -Append }
                } # if ( $pAddOns.count -gt 0 )
            } # else if ( Test-Path $lFlagFile'_BK' )
        } # if ( -not (Test-Path $lFlagFile ) )
    } else {
        $lMsg = "... $pSysID - Disabling download of AddOns Softpaqs"
        if ( Test-Path $lFlagFile ) { Move-Item $lFlagFile $lFlagFile'_BK' -Force }
    } # else if ( $pCreateFile )

    CMTraceLog -Message $lMsg -Type $TypeNorm

} # Manage_IDFlag_File

#=====================================================================================
<#
    Function AddPlatformForm
    Here we ask the user to select a new device to add to our list
#>
Function AddPlatformForm {
    [CmdletBinding()]
	param( $pGrid, $pSoftpaqList, $pRepositoryFolder, $pCommonFlag )

    $fEntryFormWidth = 400
    $fEntryFormHeigth = 400
    $fOffset = 20
    $FieldHeight = 20
    $PathFieldLength = 200

    if ( $DebugMode ) { Write-Host 'Add Entry Form' }
    $EntryForm = New-Object System.Windows.Forms.Form
    $EntryForm.MaximizeBox = $False ; $EntryForm.MinimizeBox = $False #; $EntryForm.ControlBox = $False
    $EntryForm.Text = "Find Device to Add"
    $EntryForm.Width = $fEntryFormWidth ; $EntryForm.height = 400 ; $EntryForm.Autosize = $true
    $EntryForm.StartPosition = 'CenterScreen'
    $EntryForm.Topmost = $true

    $EntryId = New-Object System.Windows.Forms.Label
    $EntryId.Text = "Name"
    $EntryId.location = New-Object System.Drawing.Point($fOffset,$fOffset) # (from left, from top)
    $EntryId.Size = New-Object System.Drawing.Size(60,20)                   # (width, height)

    $EntryModel = New-Object System.Windows.Forms.TextBox
    $EntryModel.Text = ""       # start w/INI setting
    $EntryModel.Multiline = $false 
    $EntryModel.location = New-Object System.Drawing.Point(($fOffset+70),($fOffset-4)) # (from left, from top)
    $EntryModel.Size = New-Object System.Drawing.Size($PathFieldLength,$FieldHeight)# (width, height)
    $EntryModel.ReadOnly = $False
    $EntryModel.Name = "Model Name"
    $EntryModel.add_MouseHover($ShowHelp)
    $SearchButton = New-Object System.Windows.Forms.Button
    $SearchButton.Location = New-Object System.Drawing.Point(($PathFieldLength+$fOffset+80),($fOffset-6))
    $SearchButton.Size = New-Object System.Drawing.Size(75,23)
    $SearchButton.Text = 'Search'
    $SearchButton_AddClick = {
        if ( $EntryModel.Text ) {
            $AddEntryList.Items.Clear()
            $lModels = Get-HPDeviceDetails -Like -Name $EntryModel.Text    # find all models matching string
            foreach ( $iModel in $lModels ) { [void]$AddEntryList.Items.Add($iModel.Name) }
        } # if ( $EntryModel.Text )
    } # $SearchButton_AddClick =
    $SearchButton.Add_Click( $SearchButton_AddClick )

    $lListHeight = $fEntryFormHeigth/2-60
    $AddEntryList = New-Object System.Windows.Forms.ListBox
    $AddEntryList.Name = 'Entries'
    $AddEntryList.Autosize = $false
    $AddEntryList.location = New-Object System.Drawing.Point($fOffset,60)  # (from left, from top)
    $AddEntryList.Size = New-Object System.Drawing.Size(($fEntryFormWidth-60),$lListHeight) # (width, height)
    $AddEntryList.Add_Click({
        $AddSoftpaqList.items.Clear()
        foreach ( $iName in $pSoftpaqList ) { $AddSoftpaqList.items.add($iName) }
    })
    #$AddEntryList.add_DoubleClick({ return $AddEntryList.SelectedItem })

    $SpqLabel = New-Object System.Windows.Forms.Label
    $SpqLabel.Text = "Select additional Softpaqs" 
    $SpqLabel.location = New-Object System.Drawing.Point($fOffset,($lListHeight+60)) # (from left, from top)
    $SpqLabel.Size = New-Object System.Drawing.Size(70,60)                   # (width, height)

    $AddSoftpaqList = New-Object System.Windows.Forms.ListBox
    $AddSoftpaqList.Name = 'Softpaqs'
    $AddSoftpaqList.Autosize = $false
    $AddSoftpaqList.SelectionMode = 'MultiExtended'
    $AddSoftpaqList.location = New-Object System.Drawing.Point(($fOffset+70),($lListHeight+60))  # (from left, from top)
    $AddSoftpaqList.Size = New-Object System.Drawing.Size(($fEntryFormWidth-130),($lListHeight-40)) # (width, height)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(($fEntryFormWidth-120),($fEntryFormHeigth-80))
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $EntryForm.AcceptButton = $okButton
    $EntryForm.Controls.AddRange(@($EntryId,$EntryModel,$SearchButton,$AddEntryList,$SpqLabel,$AddSoftpaqList, $okButton))
    # -------------------------------------------------------------------
    # Ask the user what model to add to the list
    # -------------------------------------------------------------------
    $lResult = $EntryForm.ShowDialog()

    if ($lResult -eq [System.Windows.Forms.DialogResult]::OK) {



        [array]$lSelectedEntry = Get-HPDeviceDetails -Like -Name $AddEntryList.SelectedItem
        #----------------------------------------------------------------------------------
        # because a model name may have several Platform IDs, add each to the list
        # ... of course, it could be just one !!
        #----------------------------------------------------------------------------------
        for ( $i=0 ; $i -lt $lSelectedEntry.count ; $i++ ) {
            [void]$pGrid.Rows.Add( @( $False, $lSelectedEntry[$i].SystemID, $lSelectedEntry[$i].Name) )

            if ( ($lSelectedSoftpaqs = $AddSoftpaqList.SelectedItems).Count -gt 0 ){
                Manage_IDFlag_File $pRepositoryFolder $lSelectedEntry[$i].SystemID $lSelectedEntry[$i].Name $lSelectedSoftpaqs $True
            } else {
                Manage_IDFlag_File $pRepositoryFolder $lSelectedEntry[$i].SystemID $lSelectedEntry[$i].Name $Null $True
            }
            Update_INIModelsList $pGrid $pRepositoryFolder $pCommonFlag
        }
        return $lSelectedEntry    
    } # if ($lResult -eq [System.Windows.Forms.DialogResult]::OK)

} # Function AddPlatformForm

########################################################################################

#=====================================================================================
<#
    Function MainForm
    This is the MAIN function with a Gui that sets things up for the user
#>
Function MainForm {
    
    #Add-Type -assembly System.Windows.Forms

    $LeftOffset = 20
    $TopOffset = 20
    $FieldHeight = 20
    $FormWidth = 900
    $FormHeight = 800

    $BackgroundColor = 'LightSteelBlue'
    
    #----------------------------------------------------------------------------------
    # ToolTips
    #----------------------------------------------------------------------------------
    $CMForm_tooltip = New-Object System.Windows.Forms.ToolTip
    $ShowHelp={
        #display popup help - each value is the 'name' of a control on the form.
        Switch ($this.name) {
            "OS_Selection"     {$tip = "What Windows 10 OS version to work with"}
            "Keep Filters"     {$tip = "Do NOT erase previous product selection filters"}
            "Continue on 404"  {$tip = "Continue Sync evern with Error 404, missing files"}
            "Individual Paths" {$tip = "Path to Head of Individual platform repositories"}
            "Common Path"      {$tip = "Path to Common/Shared platform repository"}
            "Models Table"     {$tip = "HP Models table to Sync repository(ies) to"}
            "Check All"        {$tip = "This check selects all Platforms and categories"}
            "Sync"             {$tip = "Syncronize repository for selected items from HP cloud"}
            'Refresh Grid'     {$tip = 'Reset the Grid from the INI file $HPModelsTable list'}
            'Filters'          {$tip = 'Show list of all current Repository filters'}
            'Add Model'        {$tip = 'Find and add a model to the current list in the Grid'}
        } # Switch ($this.name)
        $CMForm_tooltip.SetToolTip($this,$tip)
    } #end ShowHelp

    if ( $DebugMode ) { Write-Host 'creating Form' }
    $CM_form = New-Object System.Windows.Forms.Form
    $CM_form.Text = "HPIARepo_Downloader v$($ScriptVersion)"
    $CM_form.Width = $FormWidth
    $CM_form.height = $FormHeight
    $CM_form.Autosize = $true
    $CM_form.StartPosition = 'CenterScreen'

    #----------------------------------------------------------------------------------
    # Create OS and OS Version display fields - info from .ini file
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating OS Combo and Label' }
    $OSTextLabel = New-Object System.Windows.Forms.Label
    $OSTextLabel.Text = "Windows 10:"
    $OSTextLabel.location = New-Object System.Drawing.Point($LeftOffset,($TopOffset+4))    # (from left, from top)
    $OSTextLabel.AutoSize = $true
    $OSVERComboBox = New-Object System.Windows.Forms.ComboBox
    $OSVERComboBox.Size = New-Object System.Drawing.Size(60,$FieldHeight)                  # (width, height)
    $OSVERComboBox.Location  = New-Object System.Drawing.Point(($LeftOffset+90), ($TopOffset))
    $OSVERComboBox.DropDownStyle = "DropDownList"
    $OSVERComboBox.Name = "OS_Selection"
    $OSVERComboBox.add_MouseHover($ShowHelp)
    # populate menu list from INI file
    Foreach ($MenuItem in $v_OSVALID) {
        [void]$OSVERComboBox.Items.Add($MenuItem);
    }  
    $OSVERComboBox.SelectedItem = $v_OSVER 
    $OSVERComboBox.add_SelectedIndexChanged( {
        $v_OSVER = $OSVERComboBox.SelectedItem
        Check_PlatformsOSVersion $dataGridView $v_OSVER
        }
    ) # $OSVERComboBox.add_SelectedIndexChanged()
    
    $CM_form.Controls.AddRange(@($OSTextLabel,$OSVERComboBox))

    #----------------------------------------------------------------------------------
    # Create Keep Filters checkbox
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Keep Filters Checkbox' }
    $keepFiltersCheckbox = New-Object System.Windows.Forms.CheckBox
    $keepFiltersCheckbox.Text = 'Keep prev OS Filters'
    $keepFiltersCheckbox.Name = 'Keep Filters'
    $keepFiltersCheckbox.add_MouseHover($ShowHelp)
    $keepFiltersCheckbox.Autosize = $true
    $keepFiltersCheckbox.Location = New-Object System.Drawing.Point(($LeftOffset+160),($TopOffset+2))   # (from left, from top)

    # populate CM Udate checkbox from .INI variable setting - $v_UpdateCMPackages
    $find = "^[\$]v_KeepFilters"
    (Get-Content $IniFIleFullPath) | Foreach-Object { 
        if ($_ -match $find) { 
            if ( $_ -match '\$true' ) { 
                $keepFiltersCheckbox.Checked = $true 
            } else { 
                $keepFiltersCheckbox.Checked = $false 
            }              } 
        } # Foreach-Object

    $Script:v_KeepFilters = $keepFiltersCheckbox.Checked

    $keepFiltersCheckbox_Click = {
        $find = "^[\$]v_KeepFilters"
        if ( $keepFiltersCheckbox.checked ) {
            $Script:v_KeepFilters = $true
            CMTraceLog -Message "... Existing OS Filters will NOT be removed'" -Type $TypeWarn
        } else {
            $Script:v_KeepFilters = $false
            $keepFiltersCheckbox.Checked = $false
            CMTraceLog -Message "... Existing Filters will be removed and new filters will be created'" -Type $TypeWarn
        }
        $replace = "`$v_KeepFilters = `$$Script:v_KeepFilters"                   # set up the replacing string to either $false or $true from ini file
        Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"

    } # $keepFiltersCheckbox_Click = 

    $keepFiltersCheckbox.add_Click($keepFiltersCheckbox_Click)

    $CM_form.Controls.AddRange(@($keepFiltersCheckbox))

    #----------------------------------------------------------------------------------
    # Create Continue on Error checkbox
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Continue on Error 404 Checkbox' }
    $continueOn404Checkbox = New-Object System.Windows.Forms.CheckBox
    $continueOn404Checkbox.Text = 'Continue on Sync Error'
    $continueOn404Checkbox.Name = 'Continue on 404'
    $continueOn404Checkbox.add_MouseHover($ShowHelp)
    $continueOn404Checkbox.Autosize = $true
    $continueOn404Checkbox.Location = New-Object System.Drawing.Point(($LeftOffset+160),($TopOffset+22))   # (from left, from top)

    # populate CM Udate checkbox from .INI variable setting - $v_UpdateCMPackages
    $find = "^[\$]v_Continueon404"
    (Get-Content $IniFIleFullPath) | Foreach-Object { 
        if ($_ -match $find) { 
            if ( $_ -match '\$true' ) { 
                $continueOn404Checkbox.Checked = $true 
            } else { 
                $continueOn404Checkbox.Checked = $false 
            }              } 
        } # Foreach-Object

    $Script:v_Continueon404 = $continueOn404Checkbox.Checked

    # update .INI variable setting - $v_UpdateCMPackages 
    $continueOn404Checkbox_Click = {
        $find = "^[\$]v_Continueon404"
        if ( $continueOn404Checkbox.checked ) {
            $Script:v_Continueon404 = $true
            CMTraceLog -Message "... Will continue on Sync missing file errors'" -Type $TypeWarn
        } else {
            $Script:v_Continueon404 = $false
            $continueOn404Checkbox.Checked = $false
            CMTraceLog -Message "... Will STOP on Sync missing file errors'" -Type $TypeWarn
        }
        $replace = "`$v_Continueon404 = `$$Script:v_Continueon404"                   # set up the replacing string to either $false or $true from ini file
        Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"
    } # $continueOn404Checkbox_Click = 

    $continueOn404Checkbox.add_Click($continueOn404Checkbox_Click)

    $CM_form.Controls.AddRange(@($continueOn404Checkbox))

    #----------------------------------------------------------------------------------
    # add shared and Common repository folder fields and Radio Buttons for selection
    #----------------------------------------------------------------------------------
    # create group box to hold the repo paths

    $PathsGroupBox = New-Object System.Windows.Forms.GroupBox
    $PathsGroupBox.location = New-Object System.Drawing.Point(($LeftOffset+340),($TopOffset-10)) # (from left, from top)
    $PathsGroupBox.Size = New-Object System.Drawing.Size(($FormWidth-405),60)                    # (width, height)
    $PathsGroupBox.text = "Repository Paths - from $($IniFile):"

    #-------------------------------------------------------------------------------
    # create 'individual' radion button, label, text entry fields, and Browse button
    #-------------------------------------------------------------------------------
    $PathFieldLength = 280
    $labelWidth = 70

    if ( $DebugMode ) { Write-Host 'creating Shared radio button' }
    $SharedRadioButton = New-Object System.Windows.Forms.RadioButton
    $SharedRadioButton.Location = '10,14'
    $SharedRadioButton.Add_Click( {
            Import_IndividualRepos $dataGridView $SharePathTextField.Text $Script:HPModelsTable
        }
    ) # $SharedRadioButton.Add_Click()

    if ( $DebugMode ) { Write-Host 'creating Individual field label' }
    $SharePathLabel = New-Object System.Windows.Forms.Label
    $SharePathLabel.Text = "Root"
    #$SharePathLabel.TextAlign = "Left"    
    $SharePathLabel.location = New-Object System.Drawing.Point(($LeftOffset+15),$TopOffset) # (from left, from top)
    $SharePathLabel.Size = New-Object System.Drawing.Size($labelWidth,20)                   # (width, height)
    if ( $DebugMode ) { Write-Host 'creating Shared repo text field' }
    $SharePathTextField = New-Object System.Windows.Forms.TextBox
    $SharePathTextField.Text = "$Script:v_Root_IndividualRepoFolder"       # start w/INI setting
    $SharePathTextField.Multiline = $false 
    $SharePathTextField.location = New-Object System.Drawing.Point(($LeftOffset+100),($TopOffset-4)) # (from left, from top)
    $SharePathTextField.Size = New-Object System.Drawing.Size($PathFieldLength,$FieldHeight)# (width, height)
    $SharePathTextField.ReadOnly = $true
    $SharePathTextField.Name = "Individual Paths"
    $SharePathTextField.add_MouseHover($ShowHelp)
    
    if ( $DebugMode ) { Write-Host 'creating 'Individual' Browse button' }
    $sharedBrowse = New-Object System.Windows.Forms.Button
    $sharedBrowse.Width = 60
    $sharedBrowse.Text = 'Browse'
    $sharedBrowse.Location = New-Object System.Drawing.Point(($LeftOffset+$PathFieldLength+$labelWidth+40),($TopOffset-5))
    $sharedBrowse_Click = {
        $SharePathTextField.Text = Browse_IndividualRepos $dataGridView $SharePathTextField.Text
    } # $sharedBrowse_Click
    $sharedBrowse.add_Click($sharedBrowse_Click)

    #--------------------------------------------------------------------------
    # create radio button, 'common' label, text entry fields, and Browse button 
    #--------------------------------------------------------------------------
    $CommonRadioButton = New-Object System.Windows.Forms.RadioButton
    $CommonRadioButton.Location = '10,34'
    $CommonRadioButton.Add_Click( {
        Import_CommonRepo $dataGridView $CommonPathTextField.Text $Script:HPModelsTable
    } ) # $CommonRadioButton.Add_Click()

    if ( $DebugMode ) { Write-Host 'creating Common repo field label' }
    $CommonPathLabel = New-Object System.Windows.Forms.Label
    $CommonPathLabel.Text = "Common"
    $CommonPathLabel.location = New-Object System.Drawing.Point(($LeftOffset+15),($TopOffset+18)) # (from left, from top)
    $CommonPathLabel.Size = New-Object System.Drawing.Size($labelWidth,20)    # (width, height)
    if ( $DebugMode ) { Write-Host 'creating Common repo text field' }
    $CommonPathTextField = New-Object System.Windows.Forms.TextBox
    $CommonPathTextField.Text = "$Script:v_Root_CommonRepoFolder"
    $CommonPathTextField.Multiline = $false 
    $CommonPathTextField.location = New-Object System.Drawing.Point(($LeftOffset+100),($TopOffset+15)) # (from left, from top)
    $CommonPathTextField.Size = New-Object System.Drawing.Size($PathFieldLength,$FieldHeight)             # (width, height)
    $CommonPathTextField.ReadOnly = $true
    $CommonPathTextField.Name = "Common Path"
    $CommonPathTextField.add_MouseHover($ShowHelp)
    #$CommonPathTextField.BorderStyle = 'None'                                # 'none', 'FixedSingle', 'Fixed3D (default)'
    
    if ( $DebugMode ) { Write-Host 'creating Common Browse button' }
    $commonBrowse = New-Object System.Windows.Forms.Button
    $commonBrowse.Width = 60
    $commonBrowse.Text = 'Browse'
    $commonBrowse.Location = New-Object System.Drawing.Point(($LeftOffset+$PathFieldLength+$labelWidth+40),($TopOffset+13))
    $commonBrowse_Click = {
        Browse_CommonRepo $dataGridView $CommonPathTextField.Text
    } # $commonBrowse_Click
    $commonBrowse.add_Click($commonBrowse_Click)

    $PathsGroupBox.Controls.AddRange(@($SharePathTextField, $SharePathLabel,$SharedRadioButton, $sharedBrowse))
    $PathsGroupBox.Controls.AddRange(@($CommonPathTextField, $CommonPathLabel, $commonBrowse, $CommonRadioButton))

    $CM_form.Controls.AddRange(@($PathsGroupBox))

    #----------------------------------------------------------------------------------
    # Create Models list Checked Grid box - add 1st checkbox column
    # The ListView control allows columns to be used as fields in a row
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating DataGridView to populate with platforms' }
    $ListViewWidth = ($FormWidth-80)
    $ListViewHeight = 250
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Name = 'Models Table'
    $dataGridView.location = New-Object System.Drawing.Point(($LeftOffset-10),($TopOffset))
    $dataGridView.add_MouseHover($ShowHelp)
    $dataGridView.height = $ListViewHeight
    $dataGridView.width = $ListViewWidth
    $dataGridView.ColumnHeadersVisible = $true                   # the column names becomes row 0 in the datagrid view
    $dataGridView.ColumnHeadersHeightSizeMode = 'AutoSize'       # AutoSize, DisableResizing, EnableResizing
    $dataGridView.RowHeadersVisible = $false
    $dataGridView.SelectionMode = 'CellSelect'
    $dataGridView.AllowUserToAddRows = $False                    # Prevents the display of empty last row

    if ( $DebugMode ) {  Write-Host 'creating col 0 checkboxColumn' }
    # add column 0 (0 is 1st column)
    $CheckBoxColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $CheckBoxColumn.width = 28

    [void]$DataGridView.Columns.Add($CheckBoxColumn) 

    #----------------------------------------------------------------------------------
    # Add a CheckBox on header (to 1st col)
    # default the all checkboxes selected/checked
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Checkbox col 0 header' }
    $CheckAll=New-Object System.Windows.Forms.CheckBox
    $CheckAll.Name = 'Check All'
    $CheckAll.add_MouseHover($ShowHelp)
    $CheckAll.AutoSize=$true
    $CheckAll.Left=9
    $CheckAll.Top=6
    $CheckAll.Checked = $false

    $CheckAll_Click={
        $state = $CheckAll.Checked
        if ( $state ) {
            for ($i = 0; $i -lt $dataGridView.RowCount; $i++) {
                $dataGridView[0,$i].Value = $state  
                # 'Driver','BIOS', 'Firmware', 'Software'
                $dataGridView.Rows[$i].Cells['Driver'].Value = $state
                $dataGridView.Rows[$i].Cells['BIOS'].Value = $state
                $dataGridView.Rows[$i].Cells['Firmware'].Value = $state
                $dataGridView.Rows[$i].Cells['Software'].Value = $state
                $dataGridView.Rows[$i].Cells['AddOns'].Value = $state
            } # for ($i = 0; $i -lt $dataGridView.RowCount; $i++)
        } else {
            clear_grid $dataGridView
        }
    } # $CheckAll_Click={

    $CheckAll.add_Click($CheckAll_Click)
    
    $dataGridView.Controls.Add($CheckAll)
    
    #----------------------------------------------------------------------------------
    # add columns 1, 2 (0 is 1st column) for Platform ID and Platform Name
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'adding SysId, Model columns' }
    $dataGridView.ColumnCount = 3                                # 1st column (0) is checkbox column
    $dataGridView.Columns[1].Name = 'SysId'
    $dataGridView.Columns[1].Width = 50
    $dataGridView.Columns[1].DefaultCellStyle.Alignment = "MiddleCenter"
    $dataGridView.Columns[2].Name = 'Model'
    $dataGridView.Columns[2].Width = 230

    #################################################################

    #----------------------------------------------------------------------------------
    # Add checkbox columns for every category filter
    # from column 4 on (0 is 1st column)
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating category columns' }
    foreach ( $cat in $v_FilterCategories ) {
        $catFilter = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
        $catFilter.name = $cat
        $catFilter.width = 50
        [void]$DataGridView.Columns.Add($catFilter) 
    }
    #----------------------------------------------------------------------------------
    # add an All 'AddOns' column
    # column 7 (0 is 1st column)
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating AddOns column' }
    $CheckBoxINISoftware = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $CheckBoxINISoftware.Name = 'AddOns' 
    $CheckBoxINISoftware.width = 50
    $CheckBoxINISoftware.ThreeState = $False
    [void]$DataGridView.Columns.Add($CheckBoxINISoftware)
   
    # $CheckBoxesAll.state = Displayed, Resizable, ResizableSet, Selected, Visible

    #----------------------------------------------------------------------------------
    # add a repository path as last column
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Repo links column' }
    $LinkColumn = New-Object System.Windows.Forms.DataGridViewColumn
    $LinkColumn.Name = 'Repository'
    $LinkColumn.ReadOnly = $true

    [void]$dataGridView.Columns.Add($LinkColumn,"Repository Path")

    $dataGridView.Columns[($dataGridView.ColumnCount-1)].MinimumWidth = 300
    $dataGridView.Columns[($dataGridView.ColumnCount-1)].AutoSizeMode = 'AllCells'

    #----------------------------------------------------------------------------------
    # next clear any selection from the initial data view
    #----------------------------------------------------------------------------------
    $dataGridView.ClearSelection()
    
    #----------------------------------------------------------------------------------
    # handle ALL checkbox selections here
    # columns= 0(row selected), 1(ProdCode), 2(Prod Name), 3-6(categories, 7(AddOns), 8(repository path)
    #----------------------------------------------------------------------------------
    $CellOnClick = {

        $row = $this.currentRow.index 
        $column = $this.currentCell.ColumnIndex 
        
        # Let's see if the cell is a checkmark (type Boolean) or a text cell (which would NOT have a value of $true or $false
        # columns 1=sysId, 2=Name, 8=path (all string types)
        if ( $column -in @(0, 3, 4, 5, 6, 7) ) {

            $CellprevState = $dataGridView.rows[$row].Cells[$column].EditedFormattedValue # value 'BEFORE' click action $true/$false
            $CellNewState = !$CellprevState
            
            # here we know we are dealing with one of the checkmark/selection cells
            switch ( $column ) {
                0 {                                                               # need to reset all categories
                    if ( $CellNewState ) {
                        # default to check; Driver and ini.ps1 Software list for this model
                        # 'Driver','BIOS', 'Firmware', 'Software', then 'AddOns'
                        foreach ( $cat in $v_FilterCategories ) {                 # ... all categories          
                            $datagridview.Rows[$row].Cells[$cat].Value = $True
                        } # forech ( $cat in $v_FilterCategories )
                        $dataGridView.Rows[$row].Cells['AddOns'].Value = $true
                    } else {                               
                        foreach ( $cat in $v_FilterCategories ) {                 # ... all categories          
                            $datagridview.Rows[$row].Cells[$cat].Value = $false
                        } # forech ( $cat in $v_FilterCategories )
                        $datagridview.Rows[$row].Cells['AddOns'].Value = $false
                        $datagridview.Rows[$row].Cells[$datagridview.columnCount-1].Value = '' # ... and reset the repo path field
                    } # if ( $CellnewState )
                } # 0
                Default {   
                                                                          
                    # here to deal with clicking on a category cell or 'AddOns'

                    # if we selected the 'AddOns' let's make sure we set this as default by tapping a new file
                    # ... with platform ID as the name... 
                    # it seems as if PS can't distinguish .value from checked or not checked... always the same
                    if ( $datagridview.Rows[$row].Cells['AddOns'].State -eq 'Selected' ) {
                        if ( $v_CommonRepo ) { $lRepoPath = $CommonPathTextField.Text } else { $lRepoPath = $SharePathTextField.Text }                  
                        Manage_IDFlag_File $lRepoPath $datagridview[1,$row].value $datagridview[2,$row].value $Null $CellNewState 

                    } # if ( $datagridview.Rows[$row].Cells['AddOns'].State -eq 'Selected' )
                    if ( -not $CellNewState ) {
                        foreach ( $cat in $Script:v_FilterCategories ) {
                            $CatColumn = $datagridview.Rows[$row].Cells[$cat].ColumnIndex
                            if ( $CatColumn -eq $column ) {
                                continue                                              
                            } else {
                                # see if anouther category column is checked
                                if ( $datagridview.Rows[$row].Cells[$cat].Value ) {
                                    $CellNewState = $true
                                }
                            } # else if ( $colClicked -eq $currColumn )
                        } # foreach ( $cat in $v_FilterCategories )
                    } # if ( -not $CellNewState )

                    $datagridview.Rows[$row].Cells[0].Value = $CellNewState
                    if ( $CellNewState -eq $false ) {
                        $datagridview.Rows[$row].Cells[$datagridview.columnCount-1].Value = ''
                    }
                } # Default
            } #  switch ( $column )
        } # if ( ($dataGridView.rows[$row].Cells[$column].value -eq $true) -or ($dataGridView.rows[$row].Cells[$column].value -eq $false) )
    } # $CellOnClick = 

    $dataGridView.Add_Click($CellOnClick)

    #----------------------------------------------------------------------------------
    # Add a grouping box around the Models Grid with its name
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating GroupBox' }

    $CMModelsGroupBox = New-Object System.Windows.Forms.GroupBox
    $CMModelsGroupBox.location = New-Object System.Drawing.Point($LeftOffset,($TopOffset+50))     # (from left, from top)
    $CMModelsGroupBox.Size = New-Object System.Drawing.Size(($ListViewWidth+20),($ListViewHeight+30))       # (width, height)
    $CMModelsGroupBox.text = "HP Models / Repository Category Filters"

    $CMModelsGroupBox.Controls.AddRange(@($dataGridView))

    $CM_form.Controls.AddRange(@($CMModelsGroupBox))

    #----------------------------------------------------------------------------------
    # Add a Refresh Grid button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating a Refresh Grid button' }
    $RefreshGridButton = New-Object System.Windows.Forms.Button
    $RefreshGridButton.Location = New-Object System.Drawing.Point($LeftOffset,($FormHeight-447))    # (from left, from top)
    $RefreshGridButton.Text = 'Refresh Grid'
    $RefreshGridButton.Name = 'Refresh Grid'
    $RefreshGridButton.AutoSize=$true
    $RefreshGridButton.add_MouseHover($ShowHelp)

    $RefreshGridButton_Click={
        Empty_Grid $dataGridView
        Populate_Grid_from_INI $dataGridView $CommonPathTextField.Text $Script:HPModelsTable
    } # $RefreshGridButton_Click={

    $RefreshGridButton.add_Click($RefreshGridButton_Click)

    $CM_form.Controls.Add($RefreshGridButton)
  
    #----------------------------------------------------------------------------------
    # Add a list filters button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating a list filters button' }
    $ListFiltersdButton = New-Object System.Windows.Forms.Button
    $ListFiltersdButton.Location = New-Object System.Drawing.Point(($LeftOffset+90),($FormHeight-447))    # (from left, from top)
    $ListFiltersdButton.Text = 'Filters'
    $ListFiltersdButton.Name = 'Filters'
    $ListFiltersdButton.AutoSize=$true

    $ListFiltersdButton_Click={
        CMTraceLog -Message 'HPIA Repository Filters found...' -Type $TypeNorm
        List_Filters
    } # $ListFiltersdButton_Click={

    $ListFiltersdButton.add_Click($ListFiltersdButton_Click)

    $CM_form.Controls.Add($ListFiltersdButton)

    #----------------------------------------------------------------------------------
    # Add a Add Model button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating a Add Model button' }
    $AddModelButton = New-Object System.Windows.Forms.Button
    $AddModelButton.Width = 80
    $AddModelButton.Height = 35
    $AddModelButton.Location = New-Object System.Drawing.Point(($LeftOffset+180),($FormHeight-447))    # (from left, from top)
    $AddModelButton.Text = 'Add Model'
    $AddModelButton.Name = 'Add Model'

    $AddModelButton.add_Click( { 
        if ( $Script:v_CommonRepo ) {
            $lTempPath = $CommonPathTextField.Text } else { $lTempPath = $SharePathTextField.Text }

        if ( $null -ne ($lModelArray = AddPlatformForm $DataGridView $Script:v_Softpaqs $lTempPath $Script:v_CommonRepo ) ) {
            $Script:s_ModelAdded = $True
            CMTraceLog -Message "$($lModelArray.ProdCode):$($lModelArray.Name) added to list" -Type $TypeNorm
            CMTraceLog -Message "NOTE: INI file to be updated at next Sync" -Type $TypeWarn
        } # if ( ($lModel = Ask_AddEntry) -ne $null )
    } ) # $AddModelButton.add_Click(

    $CM_form.Controls.Add($AddModelButton)

    #----------------------------------------------------------------------------------
    # Create New Log checkbox
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating New Log checkbox' }
    $NewLogCheckbox = New-Object System.Windows.Forms.CheckBox
    $NewLogCheckbox.Text = 'New'
    $NewLogCheckbox.Autosize = $true
    $NewLogCheckbox.Location = New-Object System.Drawing.Point(($LeftOffset+280),($FormHeight-443))

    $Script:NewLog = $NewLogCheckbox.Checked

    $NewLogCheckbox_Click = {
        if ( $NewLogCheckbox.checked ) {
            $Script:NewLog = $true
        } else {
            $Script:NewLog = $false
        }
    } # $updateCMCheckbox_Click = 

    $NewLogCheckbox.add_Click($NewLogCheckbox_Click)

    $CM_form.Controls.AddRange(@($NewLogCheckbox ))

    #----------------------------------------------------------------------------------
    # Add a log file label and field
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating log file label and field' }
    $LogPathLabel = New-Object System.Windows.Forms.Label
    $LogPathLabel.Text = "Log:"
    $LogPathLabel.Location = New-Object System.Drawing.Point(($LeftOffset+335),($FormHeight-443))    # (from left, from top)
    $LogPathLabel.AutoSize=$true

    $LogPathField = New-Object System.Windows.Forms.TextBox
    $LogPathField.Text = "$Script:v_LogFile"
    $LogPathField.Multiline = $false 
    $LogPathField.location = New-Object System.Drawing.Point(($LeftOffset+370),($FormHeight-443)) # (from left, from top)
    $LogPathField.Size = New-Object System.Drawing.Size(350,$FieldHeight)                      # (width, height)
    $LogPathField.ReadOnly = $true
    $LogPathField.Name = "LogPath"
    #$LogPathField.BorderStyle = 'none'                                                  # 'none', 'FixedSingle', 'Fixed3D (default)'
    # next, move cursor to end of text in field, to see the log file name
    $LogPathField.Select($LogPathField.Text.Length,0)
    $LogPathField.ScrollToCaret()

    $CM_form.Controls.AddRange(@($LogPathLabel, $LogPathField ))

    #----------------------------------------------------------------------------------
    # Create Sync button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Sync button' }
    $buttonSync = New-Object System.Windows.Forms.Button
    $buttonSync.Width = 90
    $buttonSync.Height = 35
    $buttonSync.Text = 'Sync Repository'
    $buttonSync.Location = New-Object System.Drawing.Point(($FormWidth-130),($FormHeight-447))
    $buttonSync.Name = 'Sync'
    $buttonSync.add_MouseHover($ShowHelp)

    $buttonSync.add_click( {

        # If needed, modify INI file with newly selected v_OSVER

        if ( $Script:v_OSVER -ne $OSVERComboBox.Text ) {
            $find = "^[\$]v_OSVER"
            $replace = "`$v_OSVER = ""$($OSVERComboBox.Text)"""  
            Mod_INISetting $IniFIleFullPath $find $replace 'Changing INI file with selected v_OSVER'
        } 
        $Script:v_OSVER = $OSVERComboBox.Text                            # get selected version
        if ( $Script:v_OSVER -in $Script:v_OSVALID ) {
            # selected rows are those that have a checkmark on column 0
            # get a list of all models selected (row numbers, starting with 0)
            # ... and add each entry to an array to be used by the sync function
            $lCheckedListArray = @()
            for ($i = 0; $i -lt $dataGridView.RowCount; $i++) {
                if ($dataGridView[0,$i].Value) {
                    $lCheckedListArray += $i
                } # if  
            } # for ($i = 0; $i -lt $dataGridView.RowCount; $i++)
            if ( $Script:NewLog  ) {
                Backup_Log $Script:v_LogFile
            }
            if ( $updateCMCheckbox.checked ) {
                if ( ($Script:CMConnected = Test_CMConnection) ) {
                    if ( $DebugMode ) { CMTraceLog -Message 'Script connected to CM' -Type $TypeDebug }
                }
            } # if ( $updateCMCheckbox.checked )

            if ( $Script:v_CommonRepo ) {
                Sync_Common_Repository $dataGridView $CommonPathTextField.Text $lCheckedListArray #$Script:s_ModelAdded
                if ( $Script:s_ModelAdded ) { Update_INIModelsList $dataGridView $CommonPathTextField $True }   # $True = head of common repository
            } else {
                sync_individual_repositories $dataGridView $SharePathTextField.Text $lCheckedListArray $Script:s_ModelAdded
                if ( $Script:s_ModelAdded ) { Update_INIModelsList $dataGridView $SharePathTextField $False }   # $False = head of individual repos
            }
            $Script:s_ModelAdded = $False        # reset as the previous Sync also updated INI file
        } # if ( $Script:v_OSVER -in $Script:v_OSVALID )

    } ) # $buttonSync.add_click

    $CM_form.Controls.AddRange(@($buttonSync))

######################################################################################################
# NExt is all about MEM CM/SCCM connectivity and actions
######################################################################################################

    #----------------------------------------------------------------------------------
    # Create CM GroupBoxes
    #----------------------------------------------------------------------------------

    $CMGroupAll = New-Object System.Windows.Forms.GroupBox
    $CMGroupAll.location = New-Object System.Drawing.Point($LeftOffset,($FormHeight-405))         # (from left, from top)
    $CMGroupAll.Size = New-Object System.Drawing.Size(($FormWidth-60),60)                       # (width, height)
    $CMGroupAll.BackColor = 'lightgray'#$BackgroundColor
    $CMGroupAll.Text = 'SCCM - Disconnected'
    
    $CMGroupBox = New-Object System.Windows.Forms.GroupBox
    $CMGroupBox.location = New-Object System.Drawing.Point($LeftOffset,($TopOffset-5))         # (from left, from top)
    $CMGroupBox.Size = New-Object System.Drawing.Size(($FormWidth-650),40)                       # (width, height)
    $CMGroupBox2 = New-Object System.Windows.Forms.GroupBox
    $CMGroupBox2.location = New-Object System.Drawing.Point(($LeftOffset+260),($TopOffset-5))         # (from left, from top)
    $CMGroupBox2.Size = New-Object System.Drawing.Size(($FormWidth-320),40)                       # (width, height)

    #----------------------------------------------------------------------------------
    # Create CM Repository Packages Update button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Repo Checkbox' }
    $updateCMCheckbox = New-Object System.Windows.Forms.CheckBox
    $updateCMCheckbox.Text = 'Update Packages /'
    $updateCMCheckbox.Autosize = $true
    $updateCMCheckbox.Location = New-Object System.Drawing.Point(($LeftOffset-20),($TopOffset-5))   # (from left, from top)

    # populate CM Udate checkbox from .INI variable setting - $v_UpdateCMPackages
    $find = "^[\$]v_UpdateCMPackages"
    (Get-Content $IniFIleFullPath) | Foreach-Object { 
            if ($_ -match $find) { 
                if ( $_ -match '\$true' ) { 
                    $updateCMCheckbox.Checked = $true 
                } else { 
                    $updateCMCheckbox.Checked = $false 
                }              
            } # if ($_ -match $find)
        } # Foreach-Object

    $Script:v_UpdateCMPackages = $updateCMCheckbox.Checked

    $updateCMCheckbox_Click = {
        $Script:v_UpdateCMPackages = $updateCMCheckbox.checked
        if ( -not $Script:v_UpdateCMPackages ) {
            $Script:v_DistributeCMPackages = $false
            $CMDistributeheckbox.Checked = $false
        }
        $find = "^[\$]v_UpdateCMPackages"
        $replace = "`$v_UpdateCMPackages = `$$Script:v_UpdateCMPackages"                   # set up the replacing string to either $false or $true from ini file
        Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"

    } # $updateCMCheckbox_Click = 

    $updateCMCheckbox.add_Click($updateCMCheckbox_Click)

    #----------------------------------------------------------------------------------
    # Create CM Distribute Packages button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating DP Distro update Checkbox' }
    $CMDistributeheckbox = New-Object System.Windows.Forms.CheckBox
    $CMDistributeheckbox.Text = 'Update DPs'
    $CMDistributeheckbox.Autosize = $true
    $CMDistributeheckbox.Location = New-Object System.Drawing.Point(($LeftOffset+125),($TopOffset-5))

    # populate CM Udate checkbox from .INI variable setting - $v_DistributeCMPackages
    $find = "^[\$]v_DistributeCMPackages"
    (Get-Content $IniFIleFullPath) | Foreach-Object { 
        if ($_ -match $find) { 
            if ( $_ -match '\$true' ) { 
                $CMDistributeheckbox.Checked = $true 
            } else { 
                $CMDistributeheckbox.Checked = $false 
            }
        } # if ($_ -match $find)
        } # Foreach-Object
    $Script:v_DistributeCMPackages = $CMDistributeheckbox.Checked

    $CMDistributeheckbox_Click = {
        $Script:v_DistributeCMPackages = $CMDistributeheckbox.Checked
        $find = "^[\$]v_DistributeCMPackages"
        $replace = "`$v_DistributeCMPackages = `$$Script:v_DistributeCMPackages"                   # set up the replacing string to either $false or $true from ini file
        Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"

    } # $updateCMCheckbox_Click = 

    $CMDistributeheckbox.add_Click($CMDistributeheckbox_Click)
    $CMGroupBox.Controls.AddRange(@($updateCMCheckbox, $CMDistributeheckbox))

    #----------------------------------------------------------------------------------
    # Create update HPIA button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating update HPIA button' }
    $HPIAPathButton = New-Object System.Windows.Forms.Button
    $HPIAPathButton.Text = "Update Package"
    $HPIAPathButton.AutoSize = $true
    $HPIAPathButton.location = New-Object System.Drawing.Point(($LeftOffset-10),($TopOffset-10))   # (from left, from top)
    #$HPIAPathButton.Size = New-Object System.Drawing.Size(80,20)                              # (width, height)

    $HPIAPathButton_Click = {
        #Add-Type -AssemblyName PresentationFramework
        $lRes = [System.Windows.MessageBox]::Show("Create or Update HPIA Package in CM?","HP Image Assistant",1)    # 1 = "OKCancel" ; 4 = "YesNo"
        if ( $lRes -eq 'Ok' ) {
            $find = "^[\$]v_HPIAVersion"
            $replace = "`$v_HPIAVersion = '$($Script:v_HPIAVersion)'"       
            Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"
            $find = "^[\$]v_HPIAPath"
            $replace = "`$v_HPIAPath = '$($Script:v_HPIAPath)'"       
            Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"
            $HPIAPathField.Text = "$HPIACMPackage - $v_HPIAPath"
            CM_HPIAPackage $HPIACMPackage $v_HPIAPath $v_HPIAVersion
        }
    } # $HPIAPathButton_Click = 

    $HPIAPathButton.add_Click($HPIAPathButton_Click)

    $HPIAPathField = New-Object System.Windows.Forms.TextBox
    $HPIAPathField.Text = "$HPIACMPackage - $v_HPIAPath"
    $HPIAPathField.Multiline = $false 
    $HPIAPathField.location = New-Object System.Drawing.Point(($LeftOffset+120),($TopOffset-5)) # (from left, from top)
    $HPIAPathField.Size = New-Object System.Drawing.Size(320,$FieldHeight)                      # (width, height)
    $HPIAPathField.ReadOnly = $true
    $HPIAPathField.Name = "v_HPIAPath"
    $HPIAPathField.BorderStyle = 'none'                                                  # 'none', 'FixedSingle', 'Fixed3D (default)'

    #----------------------------------------------------------------------------------
    # Create HPIA Browse button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating HPIA Browse button' }
    $buttonBrowse = New-Object System.Windows.Forms.Button
    $buttonBrowse.Width = 60
    $buttonBrowse.Text = 'Browse'
    $buttonBrowse.Location = New-Object System.Drawing.Point(($LeftOffset+460),($TopOffset-10))
    $buttonBrowse_Click={
        $FileBrowser.InitialDirectory = $v_HPIAPath
        $FileBrowser.Title = "Browse folder for HPImageAssistant.exe"
        $FileBrowser.Filter = "exe file (*Assistant.exe) | *assistant.exe" 
        $lBrowsePath = $FileBrowser.ShowDialog()    # returns 'OK' or Cancel'
        if ( $lBrowsePath -eq 'OK' ) {
            $lLeafHPIAExeName = Split-Path $FileBrowser.FileName -leaf          
            if ( $lLeafHPIAExeName -match 'hpimageassistant.exe' ) {
                $Script:v_HPIAPath = Split-Path $FileBrowser.FileName
                $Script:v_HPIAVersion = (Get-Item $FileBrowser.FileName).versioninfo.fileversion
                $find = "^[\$]v_HPIAVersion"
                $replace = "`$v_HPIAVersion = '$($Script:v_HPIAVersion)'"       
                Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"
                $find = "^[\$]v_HPIAPath"
                $replace = "`$v_HPIAPath = '$($Script:v_HPIAPath)'"       
                Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"
                $HPIAPathField.Text = "$HPIACMPackage - $v_HPIAPath"
                CMTraceLog -Message "... HPIA Path now: ''$v_HPIAPath'' (ver.$v_HPIAVersion) - May want to update SCCM package"
            } else {
                CMTraceLog -Message "... HPIA Path [$v_HPIAPath] does not contain HPIA executable"
            }
           # write-host $lLeafHPIAName $lNewHPIAPath 
        }
    } # $buttonBrowse_Click={
    $buttonBrowse.add_Click($buttonBrowse_Click)

    $CMGroupBox2.Controls.AddRange(@($HPIAPathButton, $HPIAPathField, $buttonBrowse))
    $CMGroupAll.Controls.AddRange(@($CMGroupBox, $CMGroupBox2))
    
    $CM_form.Controls.AddRange(@($CMGroupAll))

######################################################################################################
# Done with MEM CM/SCCM actions
######################################################################################################

    #----------------------------------------------------------------------------------
    # Create Output Text Box at the bottom of the dialog
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating RichTextBox' }
    $Script:TextBox = New-Object System.Windows.Forms.RichTextBox
    $TextBox.Name = $Script:FormOutTextBox                                          # named so other functions can output to it
    $TextBox.Multiline = $true
    $TextBox.Autosize = $false
    $TextBox.ScrollBars = "Both"
    $TextBox.WordWrap = $false
    $TextBox.location = New-Object System.Drawing.Point($LeftOffset,($FormHeight-340))  # (from left, from top)
    $TextBox.Size = New-Object System.Drawing.Size(($FormWidth-60),($FormHeight/2-120)) # (width, height)

    $TextBoxFontDefault =  $TextBox.Font    # save the default font
    $TextBoxFontDefaultSize = $TextBox.Font.Size

    $CM_form.Controls.AddRange(@($TextBox))
 
    #----------------------------------------------------------------------------------
    # Add a clear TextBox button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating a clear textbox checkmark' }
    $ClearTextBox = New-Object System.Windows.Forms.Button
    $ClearTextBox.Location = New-Object System.Drawing.Point($LeftOffset,($FormHeight-47))    # (from left, from top)
    $ClearTextBox.Text = 'Clear TextBox'
    $ClearTextBox.AutoSize=$true

    $ClearTextBox_Click={
        $TextBox.Clear()
    } # $CheckWordwrap_Click={

    $ClearTextBox.add_Click($ClearTextBox_Click)

    $CM_form.Controls.Add($ClearTextBox)

    #----------------------------------------------------------------------------------
    # Create 'Debug Mode' - checkmark
    #----------------------------------------------------------------------------------
    $DebugCheckBox = New-Object System.Windows.Forms.CheckBox
    $DebugCheckBox.Text = 'Debug Mode'
    $DebugCheckBox.UseVisualStyleBackColor = $True
    $DebugCheckBox.location = New-Object System.Drawing.Point(($LeftOffset+120),($FormHeight-45))   # (from left, from top)
    $DebugCheckBox.Autosize = $true
    $DebugCheckBox.checked = $Script:DebugMode
    $DebugCheckBox.add_click( {
            if ( $DebugCheckBox.checked ) {
                $Script:DebugMode = $true
            } else {
                $Script:DebugMode = $false
            }
        }
    ) # $DebugCheckBox.add_click

    $CM_form.Controls.Add($DebugCheckBox)                    # removed CM Connect Button

    #----------------------------------------------------------------------------------
    # Add TextBox larger and smaller Font buttons
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating a TextBox smaller Font button' }
    $TextBoxFontdButtonDec = New-Object System.Windows.Forms.Button
    $TextBoxFontdButtonDec.Location = New-Object System.Drawing.Point(($LeftOffset+240),($FormHeight-47))    # (from left, from top)
    $TextBoxFontdButtonDec.Text = '< Font'
    $TextBoxFontdButtonDec.AutoSize = $true

    $TextBoxFontdButtonDec_Click={

        if ( $TextBox.Font.Size -gt 9 ) {
            $TextBox.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",($TextBox.Font.Size-2),[System.Drawing.FontStyle]::Regular)
        }
        $textbox.refresh() 

    } # $TextBoxFontdButtonDec_Click={

    $TextBoxFontdButtonDec.add_Click($TextBoxFontdButtonDec_Click)
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating a TextBox larger Font button' }
    $TextBoxFontdButtonInc = New-Object System.Windows.Forms.Button
    $TextBoxFontdButtonInc.Location = New-Object System.Drawing.Point(($LeftOffset+310),($FormHeight-47))    # (from left, from top)
    $TextBoxFontdButtonInc.Text = '> Font'
    $TextBoxFontdButtonInc.AutoSize = $true

    $TextBoxFontdButtonInc_Click={
        if ( $TextBox.Font.Size -lt 16 ) {
            $TextBox.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",($TextBox.Font.Size+2),[System.Drawing.FontStyle]::Regular)
        }
        $textbox.refresh() 
    } # $TextBoxFontdButton_Click={

    $TextBoxFontdButtonInc.add_Click($TextBoxFontdButtonInc_Click)

    $CM_form.Controls.AddRange(@($TextBoxFontdButtonDec, $TextBoxFontdButtonInc))
  
    #----------------------------------------------------------------------------------
    # Create Done/Exit Button at the bottom of the dialog
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Done/Exit Button' }
    $buttonDone = New-Object System.Windows.Forms.Button
    $buttonDone.Text = 'Exit'
    $buttonDone.Location = New-Object System.Drawing.Point(($FormWidth-120),($FormHeight-50))    # (from left, from top)

    $buttonDone.add_click( {
            $CM_form.Close()
            $CM_form.Dispose()
        }
    ) # $buttonDone.add_click
     
    $CM_form.Controls.AddRange(@($buttonDone))

    #----------------------------------------------------------------------------------
    # now, make sure we have the HP CMSL modules available to run in the script
    #----------------------------------------------------------------------------------
    Load_HPCMSLModule

    # set up the default paths for the repository based on the $Script:v_CommonRepo value from INI
    # ... find variable as long as is at start of a line (otherwise could be a comment)
    $find = "^[\$]v_CommonRepo"
    (Get-Content $IniFIleFullPath) | Foreach-Object { 
            if ($_ -match $find) {               # we found the variable
                if ( $_ -match '\$true' ) { 
                    $lCommonPath = $CommonPathTextField.Text
                    if ( ([string]::IsNullOrEmpty($lCommonPath)) -or `
                         (-not (Test-Path $lCommonPath)) ) {
                        Write-Host "Common Repository Path from INI file not found: $($lCommonPath) - Will Create" -ForegroundColor Red
                        Init_Repository $lCommonPath $True                 # $True = make it a HPIA repo
                    }   
                    Import_CommonRepo $dataGridView $lCommonPath $Script:HPModelsTable
                    $CommonRadioButton.Checked = $true                     # set the visual default from the INI setting
                    $CommonPathTextField.BackColor = $BackgroundColor
                } else { 
                    $lIndividualRoot = $SharePathTextField.Text
                    if ( [string]::IsNullOrEmpty($lIndividualRoot) ) {
                        Write-Host "Repository field is empty" -ForegroundColor Red
                    } else {
                        Import_IndividualRepos $dataGridView $lIndividualRoot $Script:HPModelsTable
                        $SharedRadioButton.Checked = $true 
                        $SharePathTextField.BackColor = $BackgroundColor
                    }
                } # else if ( $_ -match '\$true' )
            } # if ($_ -match $find)
        } # Foreach-Object

    #----------------------------------------------------------------------------------
    # Finally, show the dialog on screen
    #----------------------------------------------------------------------------------

    if ( $DebugMode ) { Write-Host 'calling ShowDialog' }
    $CM_form.ShowDialog() | Out-Null

} # Function MainForm 

########################################################################################
# --------------------------------------------------------------------------
# Start of Invocation
# --------------------------------------------------------------------------

# at this point, we are past the -h | -help runstring option... 
# ... if any runstring options left, then we need to run without the UI

# in case we need to browse for a file, create the object now
	
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }

if ( $MyInvocation.BoundParameters.count -eq 0) {
    MainForm                            # Create the GUI and take over all actions, like Report and Download
} else {
    $RunUI = $false

    if ( $PSBoundParameters.Keys.Contains('newLog') ) { Backup_Log $Script:v_LogFile } 
    CMTraceLog -Message 'HPIARepo_Downloader - BEGIN'
    CMTraceLog -Message "Script Path: $ScriptPath"
    CMTraceLog -Message "Script Name: $scriptName"
    CMTraceLog -Message "Runtime Parameters: $($MyInvocation.BoundParameters.Keys)"
    #CMTraceLog -Message "$MyInvocation.BoundParameters.Values: $($MyInvocation.BoundParameters.Values)"
    if ( $PSBoundParameters.Keys.Contains('IniFile') ) { '-iniFile: ' + $inifile }
    if ( $PSBoundParameters.Keys.Contains('RepoStyle') ) { if ( $RepoStyle -match 'Common' ) { '-RepoStyle: ' + $RepoStyle
            $v_CommonRepo = $true } else { $Script:v_CommonRepo = $false }  } 
    if ( $PSBoundParameters.Keys.Contains('Products') ) { "-Products: $($Products)" }
    if ( $PSBoundParameters.Keys.Contains('ListFilters') ) { list_filters $Script:v_CommonRepo }
    if ( $PSBoundParameters.Keys.Contains('NoIniSw') ) { '-NoIniSw' }
    if ( $PSBoundParameters.Keys.Contains('showActivityLog') ) { $showActivityLog = $true } 
    if ( $PSBoundParameters.Keys.Contains('Sync') ) { sync_repos $Script:v_CommonRepo }

} # if ( $MyInvocation.BoundParameters.count -gt 0)

########################################################################################