<#
    HP Image Assistant and Softpaq Repository Downloader
    by Dan Felman/HP Technical Consultant
        
        Loop Code: The HPModelsTable loop code, is in a separate INI.ps1 file 
            ... was created by Gary Blok's (@gwblok) post on garytown.com.
        https://garytown.com/create-hp-bios-repository-using-powershell
 
        Logging: The Log function based on code by Ryan Ephgrave (@ephingposh)
        https://www.ephingadmin.com/powershell-cmtrace-log-function/

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
    Version 1.80
        added repositories Browse buttons for both Common (shared) and Individual (rooted) folders
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

$ScriptVersion = "1.82 (August-5-2021)"

Function Show_Help {
    
} # Function Show_Help

#=====================================================================================
# manage script runtime parameters
#=====================================================================================

if ( $help ) {
    "`nRunning script without parameters opens in UI mode. This mode should be used to set up needed filters and develop Repositories"
    "`nRuntime options:"
    "`n[-help|-h] will display this text"
    "`nHPIARepo_Downloader_1.70.ps1 [[-Sync] [-ListFilters] [-inifile .\<filepath>\HPIARepo_ini.ps1] [-RepoStyle common|individual] [-products '80D4,8549,8470'] [-NoIniSw] [-ShowActivityLog]]`n"
    "... <-IniFile>, <-RepoStyle>, <-Productsts> parameters can also be positional without parameter names. In this case, every parameter counts"    
    "`nExample: HPIARepo_Downloader_1.60.ps1 .\<path>\HPIARepo_ini.ps1 Common '80D4,8549,8470'`n"    
    "-ListFilters`n`tlist repository filters in place on selected product repositories"
    "-IniFile <path to INI.ps1>`n`tthis option can be used when running from a script to set up different downloads."
    "`tIf option not give, it will default to .\HPIARepo_ini.ps1"
    "-RepoStyle {Common|Individual}`n`tthis option selects the repository style used by the downloader script"
    "`t`t'Common' - There will be a single repository used for all models - path extracted from INI.ps1 file"
    "`t`t'Individual' - Each model will have its own repository folder"
    "-Productsts '1111', '2222'`n`ta list of HP Model Product codes, as example '80D4,8549,8470'"
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

$IniFileRooted = [System.IO.Path]::IsPathRooted($IniFile)

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
#Script Vars Environment Specific loaded from INI.ps1 file

$CMConnected = $false                                # is a connection to SCCM established?
$SiteCode = $null

$AddSoftware = '.ADDSOFTWARE'                        # sub-folders where named downloaded Softpaqs will reside
$HPIAActivityLog = 'activity.log'
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

	if ($ErrorMessage -ne $null) { $Type = $TypeError }
	if ($Component -eq $null) { $Component = " " }
	if ($Type -eq $null) { $Type = $TypeNorm }

	$LogMessage = "<![LOG[$Message $ErrorMessage" + "]LOG]!><time=`"$Time`" date=`"$Date`" component=`"$Component`" context=`"`" type=`"$Type`" thread=`"`" file=`"`">"

    #$Type = 4: Debug output ($TypeDebug)
    #$Type = 10: no \newline ($TypeNoNewline)

    if ( ($Type -ne $TypeDebug) -or ( ($Type -eq $TypeDebug) -and $DebugMode) ) {
        $LogMessage | Out-File -Append -Encoding UTF8 -FilePath $Script:v_LogFile

        # Add output to GUI message box
        OutToForm $Message $Type $Script:TextBox
        
    } else {
        $lineNum = ((get-pscallstack)[0].Location -split " line ")[1]    # output: CM_HPIARepo_Downloader.ps1: line 557
        #Write-Host "$lineNum $(Get-Date -Format "HH:mm:ss") - $Message"
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
        $pMessage
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

        if ( $lCMHPIAPackage -eq $null ) {
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

    if ( $lCMRepoPackage -eq $null ) {
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

        } # if ( $pInitialize -and !(test-path "$pRepoFOlder\.Repository"))

    } # else if ( (Test-Path $pRepoFolder) -and ($pInitialize -eq $false) )

    #--------------------------------------------------------------------------------
    # intialize special folder for holding named softpaqs
    # ... this folder will hold softpaqs added outside CMSL' sync/cleanup root folder
    #--------------------------------------------------------------------------------
    $lAddSoftpaqsFolder = "$($pRepoFolder)\$($AddSoftware)"

    if ( !(Test-Path $lAddSoftpaqsFolder) -and $pInitialize ) {
        CMTraceLog -Message "... creating Add-on Softpaq Folder $lAddSoftpaqsFolder" -Type $TypeNorm
        New-Item -Path $lAddSoftpaqsFolder -ItemType directory
        if ( $DebugMode ) { CMTraceLog -Message "NO $lAddSoftpaqsFolder" -Type $TypeWarn }
    } # if ( !(Test-Path $lAddSoftpaqsFolder) )

    if ( $DebugMode ) { CMTraceLog -Message "< init_repository" -Type $TypeNorm }

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
        $lJob = start-job -Name $pJobName -scriptblock { 
            sleep $using:pSleep
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
    Function download_softpaqs_by_name
        designed to find named Softpaqs from $HPModelsTable and download those into their own repo folder
        Requires: arg1 - folder/repository 
                  prodcode - SysId of product to check for
#>
Function download_softpaqs_by_name {
    [CmdletBinding()]
	param( $pFolder,
            $pProdCode,
            $pINISWSelected )

    if ( $script:noIniSw ) {     # runstring option
        return
    }
    if ( $DebugMode ) { CMTraceLog -Message  '> download_softpaqs_by_name' }

    # let's search the HPModels list from the INI.ps1 file for Software entries

    $script:HPModelsTable  | ForEach-Object {
        if ( ($_.ProdCode -match $pProdCode) -and $pINISWSelected ) {                      # let's match the Model to download for
            $lAddSoftpaqsFolder = "$($pFolder)\$($AddSoftware)"
            Set-Location $lAddSoftpaqsFolder
            CMTraceLog -Message "... Retrieving named Softpaqs for Platform:$pProdCode" -Type $TypeNorm

            ForEach ( $Softpaq in $_.AddOns ) {
                Try {
                    Get_HPServerIPConnection 'get-softpaqlist' $HPServerIPFile 2      # check connections with 2 secs delay
                    $lListret = (Get-SoftpaqList -platform $pProdCode -os $OS -osver $Script:v_OSVER -characteristic SSM) 6>&1 
                    $lListret | Where-Object { ($_.Name -match $Softpaq) -or ($_.Id -match $Softpaq) } | ForEach-Object {
                            if ($_.SSM -eq $true) {
                                CMTraceLog -Message "      [Get Softpaq] - $($_.Id) ''$($_.Name)''" -Type $TypeNoNewLine
                                if ( Test-Path "$($_.Id).exe" ) {
                                        CMTraceLog -Message  " - Already exists. Will not re-download"  -Type $TypeWarn
                                } else {
                                    Get_HPServerIPConnection 'get-softpaq' $HPServerIPFile 2
                                    $ret = (Get-Softpaq $_.Id) 6>&1
                                    CMTraceLog -Message  "- Downloaded $($ret)"  -Type $TypeWarn 
                                    #--------------------------------------------------------------------------------
                                    # did we captured a connection to HP?
                                    Wait-Job -Name get-softpaq
                                    Receive-Job -Name get-softpaq
                                    Remove-Job -Name get-softpaq
                                    $lOut15File = "$($HPServerIPFile)_get-softpaq.TXT"
                                    # now, analyze for HP connection IP, if one found - should be in the file created
                                    if ( Test-Path $lOut15File ) {
                                        ForEach ($line in (Get-Content $lOut15File)) { CMTraceLog -Message "      $line" -Type $TypeWarn }
                                        remove-Item $lOut15File -Force        
                                    } # if ( Test-Path $usingpOutFile )
                                    #--------------------------------------------------------------------------------
                                    $ret = (Get-SoftpaqMetadataFile $_.Id) 6>&1
                                    if ( $DebugMode ) { CMTraceLog -Message  "... Get-SoftpaqMetadataFile done: $($ret)"  -Type $TypeWarn }
                                } # else if ( Test-Path "$($_.Id).exe" )
                            } else {
                                CMTraceLog -Message "      [Get Softpaq] - $($_.Id) ''$Softpaq'' not SSM compliant" -Type $TypeWarn
                            } # else if ($_.SSM -eq $true)

                        } # ForEach-Object
                        #--------------------------------------------------------------------------------
                        # see if we captured a connection to HP
                        Wait-Job -Name get-softpaqlist
                        Receive-Job -Name get-softpaqlist
                        Remove-Job -Name get-softpaqlist
                        $lOut15File = "$($HPServerIPFile)_get-softpaqlist.TXT"
                        # now, analyze for HP connection IP, if one found - should be in the file created
                        if ( Test-Path $lOut15File ) {
                            ForEach ($line in (Get-Content $lOut15File)) { CMTraceLog -Message "      $line" -Type $TypeWarn }
                            remove-Item $lOut15File -Force        
                        } # if ( Test-Path $usingpOutFile )
                        #--------------------------------------------------------------------------------
                }
                Catch {
                    $lerrMsg = "      [Get-SoftpaqList] $($lListret)... 'exception -platform:$($pProdCode) -osver:$($v_OSVER) - Filter NOT Supported!" 
                    CMTraceLog -Message $lerrMsg -Type $TypeError
                } # Catch
                    
            } # ForEach ( $Softpaq in $_.AddOns )

        } # if ( ($_.ProdCode -match $pProdCode) -and $pINISWSelected )
    } # ForEach-Object

    if ( $DebugMode ) { CMTraceLog -Message  '< download_softpaqs_by_name' }

} # Function download_softpaqs_by_name

#=====================================================================================
<#
    Function Sync_and_Cleanup_Repository
        This function will run a Sync and a Cleanup commands from HPCMSL

    expects parameter 
        - Repository folder to sync
#>
Function Sync_and_Cleanup_Repository {
    [CmdletBinding()]
	param( $pFolder )

    $lCurrentLoc = Get-Location

    if ( Test-Path $pFolder ) {

        #--------------------------------------------------------------------------------
        # update repository softpaqs with sync command and then cleanup
        #--------------------------------------------------------------------------------
        Set-Location -Path $pFolder

        #--------------------------------------------------------------------------------
        CMTraceLog -Message  '... [Sync_and_Cleanup_Repository] - please wait !!!' -Type $TypeNorm

        # but first let's start a look at what connections (port 443) are being used, wait 2 seconds
        Get_HPServerIPConnection 'invoke-repositorysync' $HPServerIPFile 2
        $lRes = (invoke-repositorysync 6>&1)
        if ( $debugMode ) { CMTraceLog -Message "... $($lRes)" -Type $TypeWarn }

        #--------------------------------------------------------------------------------
        # see if we captured a connection to HP
        Wait-Job -Name invoke-repositorysync
        Receive-Job -Name invoke-repositorysync
        Remove-Job -Name invoke-repositorysync

        $lOut15File = "$($HPServerIPFile)_invoke-repositorysync.TXT"
        # now, analyze for HP connection IP, if one found - should be 15.xx.xx.xx
        if ( Test-Path $lOut15File ) {
            ForEach ($line in (Get-Content $lOut15File)) { CMTraceLog -Message "      $line" -Type $TypeWarn }
            remove-Item $lOut15File -Force        
        } # if ( Test-Path $usingpOutFile )
        #--------------------------------------------------------------------------------

        # find what sync'd from the CMSL log file for this run
        Get_ActivityLogEntries $pFolder

        CMTraceLog -Message  '... invoking repository cleanup ' -Type $TypeNoNewline 
        $lRes = (invoke-RepositoryCleanup 6>&1)
        CMTraceLog -Message "... $($lRes)" -Type $TypeWarn

        # next, copy all softpaqs in $AddSoftware subfolder to the repository (since it got clearn up by CMSL's "Invoke-RepositoryCleanup"
        if ( -not $script:noIniSw ) {
            CMTraceLog -Message "... Copying named softpaqs (and cva files) to Repository - if selected"
            $lCmd = "robocopy.exe"
            $ldestination = "$($pFolder)"
            $lsource = "$($pFolder)\$($AddSoftware)"
            # Start robocopy -args "$source $destination $fileLIst $robocopyOptions" 
            Start $lCmd -args "$lsource $ldestination *cva"  
            Start $lCmd -args "$lsource $ldestination *exe"  
        } # if ( -not $noIniSw )

        CMTraceLog -Message  '... [Sync_and_Cleanup_Repository] - Done' -Type $TypeNorm

    } # if ( Test-Path $pFolder )

    Set-Location $lCurrentLoc

} # Function Sync_and_Cleanup_Repository

#=====================================================================================
<#
    Function Update_Repository_and_Grid
        for the selected model, 
            - remove all filters
            - add filters as selected in UI
            - invoke sync and cleanup function on the folder 
        ******* TBD: Add Platform AddOns file to .ADDSOFTWARE
#>
Function Update_Repository_and_Grid {
[CmdletBinding()]
	param( $pModelRepository,
            $pModelID,
            $pRow,
            $pAddFilters )

    $pCurrentLoc = Get-Location

    set-location $pModelRepository         # move to location of Repository to use CMSL repo commands

    if ( $Script:v_Continueon404 ) {
        Set-RepositoryConfiguration -Setting OnRemoteFileNotFound -Value LogAndContinue
    } else {
        Set-RepositoryConfiguration -Setting OnRemoteFileNotFound -Value Fail
    }
    if ( $pAddFilters ) {
    
        if ( $Script:v_KeepFilters ) {
            CMTraceLog -Message  "... Keeping existing filters in: $($pModelRepository) -platform $($pModelID)" -Type $TypeNorm
        } else {
            CMTraceLog -Message  "... Removing existing filters in: $($pModelRepository) -platform $($pModelID)" -Type $TypeNorm
            $lres = (Remove-RepositoryFilter -platform $pModelID -yes 6>&1)      
            if ( $debugMode ) { CMTraceLog -Message "... removed filter: $($lres)" -Type $TypeWarn }
        } # else if ( $Script:v_KeepFilters )
        #--------------------------------------------------------------------------------
        # now update filters - every category checked for the current model'
        #--------------------------------------------------------------------------------

        if ( $DebugMode ) { CMTraceLog -Message "... adding category filters" -Type $TypeDebug }

        foreach ( $cat in $v_FilterCategories ) {
            if ( $datagridview.Rows[$pRow].Cells[$cat].Value ) {
                CMTraceLog -Message  "... adding filter: -platform $($pModelID) -os $OS -osver $v_OSVER -category $($cat) -characteristic ssm" -Type $TypeNoNewline
                $lRes = (Add-RepositoryFilter -platform $pModelID -os win10 -osver $v_OSVER -category $cat -characteristic ssm 6>&1)
                CMTraceLog -Message $lRes -Type $TypeWarn 
            }
        } # foreach ( $cat in $v_FilterCategories )

        # if we are maintaining AddOns Softpaqs, see it in the checkbox (column 7)
        $lAddOnsRepoFile = $pModelRepository+'\'+$addsoftware+'\'+$pModelID 
        if ( $datagridview[7,$pRow].Value ) {
            if (Test-Path $lAddOnsRepoFile) {

            } # if ( Test-Path $lAddOnRepoFile )
        } # if ( $datagridview[7,$pRow].Value )
<#
        $lAddOnsRepoFile = $pModelRepository+'\'+$addsoftware+'\'+$pModelID 
                
                    if ( -not ($pDataGrid.Rows[$i].Cells['AddOns'].Value) ) { # expose only once/platform
                        $pDataGrid.Rows[$i].Cells['AddOns'].Value = $True
                        [array]$lAddOns = Get-Content $lAddOnsRepoFile
                        CMTraceLog -Message "...Additional Softpaqs: $lAddOns" -Type $TypeWarn  
                    } #                    
                } # if ( Test-Path $lAddOnRepoFile )
#>
        #--------------------------------------------------------------------------------
        # update repository path field for this model in the grid (last col is links col)
        #--------------------------------------------------------------------------------
        $datagridview[($dataGridView.ColumnCount-1),$pRow].Value = $pModelRepository

    } else {

        if ( $DebugMode ) { CMTraceLog -Message "... Removing filters, -platform $($pModelID)" -Type $TypeDebug }

        $datagridview[($dataGridView.ColumnCount-1),$pRow].Value = ""

        # remove filters here
        # see if there are filters for this product in place
        $lProdFilters = (get-repositoryinfo).Filters                # from current location (repo) folder

        foreach ( $filter in $lProdFilters ) {
            <# ( objects returned by (Get-RepositoryInfo).filters example )
                platform        : 83B2
                operatingSystem : win10:1909
                category        : Firmware
                releaseType     : *
                characteristic  : ssm
            #>
            if ( $filter.platform -match $pModelID ) {
                CMTraceLog -Message "... unselected platform $($filter.platform) matched - filters removed" -Type $TypeWarn 
                $lres = (Remove-RepositoryFilter -platform $pModelID -yes 6>&1)      
                if ( $debugMode ) { CMTraceLog -Message "... removed filter: $($lres)" -Type $TypeWarn }
            }
        } # foreach ( $filter in $lProdFilters )

    } # else if ( $pAddFilters ) {

    Set-Location -Path $pCurrentLoc

} # Update_Repository_and_Grid

#=====================================================================================
<#
    Function Sync_Repos
    This function is used for command-line execution
#>
Function Sync_Repos {

    $lCurrentSetLoc = Get-Location

    #--------------------------------------------------------------------------------
    # find out if the share exists, if not, just return
    #--------------------------------------------------------------------------------

    if ( $Script:v_CommonRepo ) {
            
        # basic check to confirm the repository was configured for HPIA
        if ( !(Test-Path "$($Script:v_Root_CommonRepoFolder)\.Repository") ) {
            CMTraceLog -Message  "... Common Repository Folder selected, not initialized" -Type $TypeNorm
            return
        } 
        set-location $Script:v_Root_CommonRepoFolder                 
        Sync_and_Cleanup_Repository $Script:v_Root_CommonRepoFolder  

    } else {

        # basic check to confirm the repository exists that hosts individual repos
        if ( Test-Path $Script:v_Root_IndividualRepoFolder ) {
            # let's traverse every product Repository folder
            $lProdFolders = Get-ChildItem -Path $Script:v_Root_IndividualRepoFolder | where {($_.psiscontainer)}

            foreach ( $lprodName in $lProdFolders ) {
                $lCurrentPath = "$($Script:v_Root_IndividualRepoFolder)\$($lprodName.name)"
                set-location $lCurrentPath
                Sync_and_Cleanup_Repository $lCurrentPath  
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
            - invoke sync and cleanup function on the folder 
#>
Function sync_individual_repositories {
[CmdletBinding()]
	param( $pModelsList,                                             # array of row lines that are checked
            $pCheckedItemsList)                                      # array of rows selected
    
    CMTraceLog -Message "> sync_individual_repositories - START" -Type $TypeNorm
    if ( $DebugMode ) { CMTraceLog -Message "... for OS version: $($v_OSVER)" -Type $TypeDebug }

    #--------------------------------------------------------------------------------
    # decide if we work on single repo or individual repos
    #--------------------------------------------------------------------------------

    $lMainRepo =  $Script:v_Root_IndividualRepoFolder            
    init_repository $lMainRepo $false                            # make sure Main repo folder exists, do NOT make it a repository

    #--------------------------------------------------------------------------------
    if ( $DebugMode ) { CMTraceLog -Message "... stepping through all selected models" -Type $TypeDebug }

    # go through every selected HP Model in the list

    for ( $i = 0; $i -lt $pModelsList.RowCount; $i++ ) {
        
        $lModelId = $pModelsList[1,$i].Value                      # column 1 has the Model/Prod ID
        $lModelName = $pModelsList[2,$i].Value                    # column 2 has the Model name

        # if model entry is checked, we need to create a repository, but ONLY if $v_CommonRepo = $false
        if ( $i -in $pCheckedItemsList ) {

            CMTraceLog -Message "`n--- Updating model: $lModelName"

            $lTempRepoFolder = "$($lMainRepo)\$($lModelName)"     # this is the repo folder for this model
            init_repository $lTempRepoFolder $true
            
            Update_Repository_and_Grid $lTempRepoFolder $lModelId $i $true    # $true means 'add filters'
            
            #--------------------------------------------------------------------------------
            # update specific Softpaqs as defined in INI file for the current model
            #--------------------------------------------------------------------------------
            download_softpaqs_by_name  $lTempRepoFolder $lModelID $pModelsList.Rows[$i].Cells['AddOns'].Value

            #--------------------------------------------------------------------------------
            # now sync up and cleanup this repository
            #--------------------------------------------------------------------------------
            Sync_and_Cleanup_Repository $lTempRepoFolder 

            #--------------------------------------------------------------------------------
            # update SCCM Repository package, if user allows
            #--------------------------------------------------------------------------------
            if ( $Script:v_UpdateCMPackages ) {
                CM_RepoUpdate $lModelName $lModelId $lTempRepoFolder
            }
            
        } # if ( $i -in $pCheckedItemsList )

    } # for ( $i = 0; $i -lt $pModelsList.RowCount; $i++ )

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
	param( $pGridList,                                      # the grid of platforms
            $pCommonRepo,
            $pCheckedItemsList )                            # array of rows selected
    
    $pCurrentLoc = Get-Location
    CMTraceLog -Message "> Sync_Common_Repository - START" -Type $TypeNorm
    if ( $DebugMode ) { CMTraceLog -Message "... for OS version: $($Script:v_OSVER)" -Type $TypeDebug }

    init_repository $pCommonRepo $true                      # make sure Main repo folder exists, or create it - no init
    CMTraceLog -Message  "... Common repository selected: $($pCommonRepo)" -Type $TypeNorm

    $lAskedToRemoveFilters = $false

    #--------------------------------------------------------------------------------
    if ( $DebugMode ) { CMTraceLog -Message "... stepping through selected models" -Type $TypeDebug }

    # loop through every Model in the list

    for ( $item = 0; $item -lt $pGridList.RowCount; $item++ ) {
        
        $lModelId = $pGridList[1,$item].Value                # column 1 has the Model/Prod ID
        $lModelName = $pGridList[2,$item].Value              # column 2 has the Model name
        $lAddOns = $pGridList[7,$item].Value                 # column 7 has the AddOns checkmark

        # if model entry row is selected, then work on the platform
        if ( $item -in $pCheckedItemsList ) {

            CMTraceLog -Message "--- Updating model: $lModelName"

            # update repo filters and show in grid
            Update_Repository_and_Grid $pCommonRepo $lModelId $item $true    # $true means 'add filters'
            #--------------------------------------------------------------------------------
            # update SCCM Repository package, if user decided with the UI checkmark
            #--------------------------------------------------------------------------------
            if ( $Script:v_UpdateCMPackages ) {
                CM_RepoUpdate $lModelName $lModelId $pCommonRepo
            }
            #--------------------------------------------------------------------------------
            # update specific Softpaqs as defined in INI file for the current model
            #--------------------------------------------------------------------------------
            download_softpaqs_by_name  $pCommonRepo $lModelID $pGridList.Rows[$item].Cells['AddOns'].Value

        } else {
            # ask ONCE (this entry happens on every model entry)
            if ( -not $lAskedToRemoveFilters ) {
                $lResponse = [System.Windows.MessageBox]::Show("Remove previous filters for products not selected?","Common Repository",4)    # 1 = "OKCancel" ; 4 = "YesNo"
                $lAskedToRemoveFilters = $true
            }
            if ( $lResponse -eq 'Yes' ) {
                Update_Repository_and_Grid $pCommonRepo $lModelId $item $false   # $false means 'do not add filters, just remove them'
            }

        } # if ( $item -in $pCheckedItemsList )

    } # for ( $item = 0; $item -lt $pGridList.RowCount; $item++ )

    # we are done checking every model for filters, so now do a Softpaq Sync and cleanup

    Sync_and_Cleanup_Repository $pCommonRepo

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
        set-location $Script:v_Root_IndividualRepoFolder | where {($_.psiscontainer)}

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
	param( $pDataGrid,                              # array of row lines that are checked
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
        CMTraceLog -Message "Refreshing Grid from Common Repository: ''$pCommonFolder''" -Type $TypeNorm
        clear_grid $pDataGrid
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
        for ( $i = 0; $i -lt $pDataGrid.RowCount; $i++ ) {
            $lPlatform = $filter.platform
            if ( ($lPlatform -eq $pDataGrid[1,$i].value) -and $pRefreshGrid ) {
                # we matched the row/SysId with the Filter

                # ... so let's add each category in the filter to the Model in the GUI
                foreach ( $cat in  ($filter.category.split(' ')) ) {
                    $pDataGrid.Rows[$i].Cells[$cat].Value = $true
                }
                $pDataGrid[0,$i].Value = $true   # check the selection column 0
                $pDataGrid[($pDataGrid.ColumnCount-1),$i].Value = $pCommonFolder # ... and the Repository Path
                 
                # if we are maintaining AddOns Softpaqs, show it in the checkbox
                $lAddOnsRepoFile = $pCommonFolder+'\'+$addsoftware+'\'+$lPlatform 
                if (Test-Path $lAddOnsRepoFile) {
                    if ( -not ($pDataGrid.Rows[$i].Cells['AddOns'].Value) ) { # expose only once/platform
                        $pDataGrid.Rows[$i].Cells['AddOns'].Value = $True
                        [array]$lAddOns = Get-Content $lAddOnsRepoFile
                        CMTraceLog -Message "...Additional Softpaqs: $lAddOns" -Type $TypeWarn  
                    } #                    
                } # if ( Test-Path $lAddOnRepoFile )
            } # if ( ($lPlatform -eq $pDataGrid[1,$i].value) -and $pRefreshGrid )
        } # for ( $i = 0; $i -lt $pDataGrid.RowCount; $i++ )
    } # foreach ( $filter in $lProdFilters )

    [void]$pDataGrid.Refresh
    CMTraceLog -Message "Refreshing Grid ...DONE" -Type $TypeSuccess

    if ( $DebugMode ) { CMTraceLog -Message '< Get_CommonRepofilters]' -Type $TypeNorm }
    
} # Function Get_CommonRepofilters

#=====================================================================================
<#
    Function Get_IndividualRepofilters
        Retrieves category filters from the share for each selected model
        ... and populates the Grid appropriately
        Parameters:
            $pDataGrid                          The models grid in the GUI
            $pRepoLocation                      Where to start looking for repositories
            $pRefreshGrid                       $True to refresh the GUI grid, $false to just list filters
#>
Function Get_IndividualRepofilters {
    [CmdletBinding()]
	param( $pDataGrid,                                  # array of row lines that are checked
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
    for ( $i = 0; $i -lt $pDataGrid.RowCount; $i++ ) {

        $lModelId = $pDataGrid[1,$i].Value                     # column 1 has the Model/Prod ID
        $lModelName = $pDataGrid[2,$i].Value                   # column 2 has the Model name
        $lTempRepoFolder = "$($pRepoLocation)\$($lModelName)"  # this is the repo folder for this model
        
        if ( Test-Path $lTempRepoFolder ) {
            set-location $lTempRepoFolder                      # move to location of Repository to use CMSL repo commands
                                    
            $lProdFilters = (get-repositoryinfo).Filters

            foreach ( $platform in $lProdFilters ) {
                if ( $pRefreshGrid ) {                    
                    CMTraceLog -Message "... populating filter categories ''$($lModelName)''" -Type $TypeDebug
                    foreach ( $cat in  ($lProdFilters.category.split(' ')) ) {
                        $pDataGrid.Rows[$i].Cells[$cat].Value = $true
                    }
                    #--------------------------------------------------------------------------------
                    # show repository path for this model (model is checked) (col 8 is links col)
                    #--------------------------------------------------------------------------------
                    $pDataGrid[0,$i].Value = $true
                    $pDataGrid[($dataGridView.ColumnCount-1),$i].Value = $lTempRepoFolder

                    # if we are maintaining AddOns Softpaqs, show it in the checkbox
                    $lAddOnRepoFile = $lTempRepoFolder+'\'+$addsoftware+'\'+$lProdFilters.platform   
                    if ( Test-Path $lAddOnRepoFile ) { 
                        $pDataGrid.Rows[$i].Cells['AddOns'].Value = $True
                        $lMsg = "...Additional Softpaqs Enabled for Platform ''$($lProdFilters.platform)'': "+$Script:HPModelsTable[$i].AddOns
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
    Function get_filters
        Retrieves category filters from the share for each selected model
        ... and populates the Grid appropriately
#>
Function get_filters {
    [CmdletBinding()]
	param( $pDataGrid,                            # array of row lines that are checked
        $pRefreshGrid )                           # if $true, refresh grid with filters from repository

    $pCurrentLoc = Get-Location
    
    if ( $DebugMode ) { CMTraceLog -Message '> get_filters' -Type $TypeNorm }

    if ( $pRefreshGrid ) {
        clear_grid $pDataGrid
    }
    #--------------------------------------------------------------------------------
    # find out if the share exists (also check for empty path), if not, just return
    #--------------------------------------------------------------------------------

    $lPathIsValid = $False
    
    if ( $Script:v_CommonRepo ) {
        if ( -not ([string]::IsNullOrWhiteSpace($Script:v_Root_CommonRepoFolder)) ) {
            $lPathIsValid = $True
            Get_CommonRepofilters $pDataGrid $Script:v_Root_CommonRepoFolder $pRefreshGrid
        }
    } else {
        if ( -not ([string]::IsNullOrWhiteSpace($Script:v_Root_IndividualRepoFolder)) ) {
            $lPathIsValid = $True
            Get_IndividualRepofilters $pDataGrid $Script:v_Root_IndividualRepoFolder $pRefreshGrid 
        }
    } # else if ( $Script:v_CommonRepo )

    if ( -not $lPathIsValid ) {
        if ( $DebugMode ) { CMTraceLog -Message 'Main Repo Host Folder ''$($lMainRepo)'' does NOT exist, or NOT valid' -Type $TypeDebug }
    }

    [void]$pDataGrid.Refresh

    if ( $DebugMode ) { CMTraceLog -Message '< get_filters]' -Type $TypeNorm }
    
} # Function get_filters

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
        Paramaters: $pDataGrid = pointer to model grid in UI
                    $pCurrentRepository
#>
Function Browse_IndividualRepos {
    [CmdletBinding()]
	param( $pDataGrid, $pCurrentRepository )  

    # let's find the repository to import
    $lbrowse = New-Object System.Windows.Forms.FolderBrowserDialog
    $lbrowse.SelectedPath = $pCurrentRepository         
    $lbrowse.Description = "Select an HPIA Head Repository Folder"
    $lbrowse.ShowNewFolderButton = $true
                                  
    if ( $lbrowse.ShowDialog() -eq "OK" ) {
        $lRepository = $lbrowse.SelectedPath
        CMTraceLog -Message  "... clearing Model list" -Type $TypeNorm
        Empty_Grid $pDataGrid
        CMTraceLog -Message  "... moving to location: $($lRepository)" -Type $TypeNorm
        Set-Location $lRepository
        $lDirectories = (ls -Directory)     # Let's search for repositories at this location
        # add a row in the DataGrid for every repository we find
        foreach ( $lFolder in $lDirectories ) {
            $lRepoFolder = $lRepository+'\'+$lFolder
            Set-Location $lRepoFolder
            Try {
                $lRepoFilters = (Get-RepositoryInfo).Filters
                CMTraceLog -Message  "... Repository Found ($($lRepoFolder))" -Type $TypeNorm
                CMTraceLog -Message  "... adding platform to grid" -Type $TypeNorm
                # ... and parse to find tge platform SysID in the repository
                [array]$lRepoPlatforms = $lRepoFilters.platform
                [void]$pDataGrid.Rows.Add(@( $true, $lRepoPlatforms, $lFolder ))
            } Catch {
                CMTraceLog -Message  "... $($lRepoFolder) is NOT a Repository" -Type $TypeNorm
            } # Catch
        } # foreach ( $lFolder in $lDirectories )
        
        CMTraceLog -Message  "... Getting Filters from repositories"
        Get_IndividualRepofilters $pDataGrid $lRepository $True
        Update_UIandINI $lRepository $False              # $False = using individual repositories
        Update_INIModels $pDataGrid $lRepository $False  # $False = treat as head of individual repositories
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
        Paramaters: $ppDataGridList = pointer to model grid in UI
                    $pImportCurrentOnly = $true used during radio button selection, to use current HPModelsTable info 
                               $false is to find/browse for individual repos as head folder
        returns: object w/Folder picked for repository, or $null if user cancelled
                Folder return value contains True/Fail entry [0] + folder name [1]
#>
Function Import_IndividualRepos {
    [CmdletBinding()]
	param( $pDataGrid, $pCurrentRepository )                                

    if ( $DebugMode ) { CMTraceLog -Message '> Import_IndividualRepos' -Type $TypeNorm }

    if ( ([string]::IsNullOrEmpty($pCurrentRepository)) -or (-not (Test-Path $pCurrentRepository)) ) {
        CMTraceLog -Message  "... No Repository to import" -Type $TypeWarn
        return
    }
    CMTraceLog -Message  "... clearing Model list" -Type $TypeNorm
    Empty_Grid $pDataGrid
    CMTraceLog -Message  "... moving to location: $($pCurrentRepository)" -Type $TypeNorm
    
    Set-Location $pCurrentRepository

    # Let's search for repositories at this location
    $lDirectories = (ls -Directory)

    if ( $lDirectories.count -eq 0 ) {
        Populate_Grid_from_INI $pDataGrid $pCurrentRepository $Script:HPModelsTable
    } else {
        foreach ( $lProdName in $lDirectories ) {
            $lRepoFolderFullPath = $pCurrentRepository+'\'+$lProdName
            Set-Location $lRepoFolderFullPath
            Try {
                $lRepoFilters = (Get-RepositoryInfo).Filters
                CMTraceLog -Message  "... Repository Found ($($lRepoFolderFullPath))" -Type $TypeNorm
                CMTraceLog -Message  "... reading repository filters" -Type $TypeNorm
                [void]$pDataGrid.Rows.Add(@( $true, $lRepoFilters.platform, $lProdName ))
            } Catch {
                CMTraceLog -Message  "... NOT a Repository: ($($lRepoFolderFullPath))" -Type $TypeNorm
            } # Catch
        } # foreach ( $lProdName in $lDirectories )
        CMTraceLog -Message  "... Getting Filters from repositories"
        Get_IndividualRepofilters $pDataGrid $pCurrentRepository $True
        CMTraceLog -Message  "... Updating UI and INI Path"
        Update_UIandINI $pCurrentRepository $False
        CMTraceLog -Message "Import Individual Repositories Done ($($lRepository))" -Type $TypeSuccess
    } # else if ( $lDirectories.count -eq 0 )

    if ( $DebugMode ) { CMTraceLog -Message '< Import_IndividualRepos]' -Type $TypeNorm }

} # Function Import_IndividualRepos

#=====================================================================================
<#
    Function Browse_CommonRepo
        Browse to find existing or create new repository
        Paramaters: $pDataGrid = pointer to model grid in UI
                    $pCurrentRepository
        returns: object w/Folder picked for repository, or $null if user cancelled
                Folder return value contains True/Fail entry [0] + folder name [1]
#>
Function Browse_CommonRepo {
    [CmdletBinding()]
	param( $pDataGrid, $pCurrentRepository )                                

    # let's find the repository to import or use the current

    $lbrowse = New-Object System.Windows.Forms.FolderBrowserDialog
    $lbrowse.SelectedPath = $pCurrentRepository        # start with the repo listed in the INI file
    $lbrowse.Description = "Select an HPIA Common/Shared Repository"
    $lbrowse.ShowNewFolderButton = $true

    if ( $lbrowse.ShowDialog() -eq "OK" ) {
        $lRepository = $lbrowse.SelectedPath
        Empty_Grid $pDataGrid
        #--------------------------------------------------------------------------------
        # find out if the share exists and has the context for HPIA, if not, just return
        #--------------------------------------------------------------------------------
        Try {
            Set-Location $lRepository
            $lProdFilters | Out-Host
            $lProdFilters = (Get-RepositoryInfo).Filters
             # ... populate grid with platform SysIDs found in the repository
            [array]$lRepoPlatforms = $lProdFilters.platform | Get-Unique
            for ( $i = 0 ; $i -lt $lRepoPlatforms.count ; $i++) {
                [array]$lProdName = Get-HPDeviceDetails -Platform $lRepoPlatforms[$i]
                #$row = @( $false, $lRepoPlatforms[$i], $lProdName[0].Name )  
                [void]$pDataGrid.Rows.Add(@( $false, $lRepoPlatforms[$i], $lProdName[0].Name ))
            } # for ( $i = 0 ; $i -lt $lRepoPlatforms.count ; $i++)
            Get_CommonRepofilters $pDataGrid $lRepository $True
            Update_UIandINI $lRepository $True
            Update_INIModels $pDataGrid $lRepository $True  # $True = this is a common repository
            CMTraceLog -Message "... found HPIA repository $($lRepository)" -Type $TypeNorm
        } Catch {
            # lets initialize repository but maintain existing models in the grid
            # ... in case we want to use some of them in the new repo
            init_repository $lRepository $true
            clear_grid $pDataGrid
            CMTraceLog -Message "... $($lRepository) initialized " -Type $TypeNorm
        } # catch
        CMTraceLog -Message "Browse Common Repository Done ($($lRepository))" -Type $TypeSuccess
    } else {
        $lRepository = $pCurrentRepository
    } # else if ( $lbrowse.ShowDialog() -eq "OK" )

    return $lRepository

} # Function Browse_CommonRepo

#=====================================================================================
<#
    Function Import_CommonRepo
        If this is an existing repository with filters, show the contents in the grid
            and update INI file about imported repository

        else, populate grid from INI's platform list

        Paramaters: $pDataGrid = pointer to model grid in UI
                    $pCurrentRepository
                    $pModelsTable               --- current models table to use if nothing in repository
#>
Function Import_CommonRepo {
    [CmdletBinding()]
	param( $pGrid, $pCurrentRepository, $pModelsTable )                                

    if ( $DebugMode ) { CMTraceLog -Message '> Import_CommonRepo' -Type $TypeNorm }

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
        CMTraceLog -Message  "... HPIA Repository Found ($($lRepository))" -Type $TypeNorm
        CMTraceLog -Message  "... clearing Model list" -Type $TypeNorm
        Empty_Grid $pGrid
        CMTraceLog -Message  "... finding filters" -Type $TypeNorm

        if ( $lProdFilters.count -eq 0 ) {
            CMTraceLog -Message  "... no filters found, will fill grid from INI contents" -Type $TypeNorm
            Populate_Grid_from_INI $pGrid $pCurrentRepository $pModelsTable
        } else {
            # let's update the grid from what's in the repository ...
            # ... get the list of platform SysIDs in the repository
            [array]$lRepoPlatforms = $lProdFilters.platform | Get-Unique
            # next, add each product to the grid, to then populate with the filters
            CMTraceLog -Message  "... Adding platforms from repository"
            for ( $i = 0 ; $i -lt $lRepoPlatforms.count ; $i++) {
                # ... treat as array, as some systems have common platform IDs, so we'll pick the 1st entry
                [array]$lProdName = Get-HPDeviceDetails -Platform $lRepoPlatforms[$i]
                [void]$pGrid.Rows.Add(@( $false, $lRepoPlatforms[$i], $lProdName[0].Name ))
            } # for ( $i = 0 ; $i -lt $lRepoPlatforms.count ; $i++)
            CMTraceLog -Message  "... Reading Filters from repository" -Type $TypeNorm
            Get_CommonRepofilters $pGrid $pCurrentRepository $True
            CMTraceLog -Message  "... Updating UI and INI file" -Type $TypeNorm
            Update_UIandINI $pCurrentRepository $True
            CMTraceLog -Message  "... Updating models in INI file" -Type $TypeNorm
            Update_INIModels $pGrid $pCurrentRepository $True  # $False means treat as head of individual repositories
        } # if ( $lProdFilters.count -gt 0 )
        CMTraceLog -Message "Import Repository Done ($($lRepository))" -Type $TypeSuccess
    } Catch {
        CMTraceLog -Message  "... Repository Folder ($($lRepository)) not initialized for HPIA" -Type $TypeWarn
    } # Catch
   
    if ( $DebugMode ) { CMTraceLog -Message '< Import_CommonRepo]' -Type $TypeNorm }

} # Function Import_CommonRepo

#=====================================================================================
<#
    Function Check_PlatformsOSVersion
    here we check the OS version, if supported by any platform checked in list
#>
Function Check_PlatformsOSVersion  {
    [CmdletBinding()]
	param( $pDataGridList,
        $pOSVersion )

    if ( $DebugMode ) { CMTraceLog -Message '> Check_PlatformsOSVersion]' -Type $TypeNorm }
    
    CMTraceLog -Message "Checking support for Win10/$($pOSVersion) for selected platforms" -Type $TypeNorm

    # search thru the table entries for checked items, and see if each product
    # has support for the selected OS version

    for ( $i = 0; $i -lt $pDataGridList.RowCount; $i++ ) {

        if ( $dataGridView[0,$i].Value ) {
            $lPlatform = $pDataGridList[1,$i].Value
            $lPlatformName = $pDataGridList[2,$i].Value

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
                    CMTraceLog -Message "-- OS supported: $($lPlatform)/$($lPlatformName)" -Type $TypeNorm
                    $v_OSVER = $pOSVersion    # reset to the other version
                    write-host 'new OS ver: '$v_OSVER
                } else {
                    CMTraceLog -Message "-- OS NOT supported by: $($lPlatform)/$($lPlatformName)" -Type $TypeError
                }
            } # else if ( $pOSVersion -in ($lOSList).OperatingSystemRelease )
        } # if ( $dataGridView[0,$i].Value )  

    } # for ( $i = 0; $i -lt $pDataGridList.RowCount; $i++ )
    
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
        $pText - the comment for output
#>
Function Mod_INISetting {
    [CmdletBinding()]
	param( $pFile, $pfind, $pReplace, $pText )

    (Get-Content $pFile) | Foreach-Object {if ($_ -match $pfind) {$preplace} else {$_}} | Set-Content $pFile
    
    CMTraceLog -Message $pText -Type $TypeNorm
    
} # Mod_INISetting

#=====================================================================================
<#
    Function Update_INIModels
        Create text file with $HPModelsTable array from current checked grid entries
        Parameters
            $pDataGrid
            $pRepositoryFolder
            $pCommonRepoFlag
#>
Function Update_INIModels {
    [CmdletBinding()]
	param( $pDataGrid,
        $pRepositoryFolder,
        $pCommonRepoFlag )

    #$lModelsListFile = (Split-Path $IniFIleFullPath -Parent) +'\HPIARepo_ini - Copy.ps1'
    $lModelsListINIFile = $Script:IniFIleFullPath
    if ( $DebugMode ) { CMTraceLog -Message '> Update_INIModels' -Type $TypeNorm }

    # -------------------------------------------------------------------
    # create list of models from the grid - assumes GUI grid is populated
    # -------------------------------------------------------------------
    $lModelsList = @()
    for ( $i = 0; $i -lt $pDataGrid.RowCount; $i++ ) {
        $lModelId = $pDataGrid[1,$i].Value             # column 1 has the Platform ID
        $lModelName = $pDataGrid[2,$i].Value           # column 2 has the Model name
        # add AddOns entries
        $lAddOnsFlagFile = $pRepositoryFolder+'\'+$addsoftware+'\'+$lModelId
        if (Test-Path $lAddOnsFlagFile) {
            # create an array string of softpaq names to add to the platform entry in the models list
            [array]$lAddOns = Get-Content $lAddOnsFlagFile
            $lAddOnsString = "'$($lAddOns[0])'"
            for ($iAdds=1; $iAdds -lt $lAddOns.Count; $iAdds++) { $lAddOnsString += ", '$($lAddOns[$iAdds])'" }
            $lAddModel = "`n`t@{ ProdCode = '$($lModelId)'; Model = '$($lModelName)' ; AddOns = $($lAddOnsString) }"
        } else {
            $lAddModel = "`n`t@{ ProdCode = '$($lModelId)'; Model = '$($lModelName)' }"
        }
        $lModelsList += $lAddModel                      # add the entry t othe model list
    } # for ( $i = 0; $i -lt $pDataGrid.RowCount; $i++ )
    CMTraceLog -Message '... Created HP Models List' -Type $TypeNorm

    if ( Test-Path $lModelsListINIFile ) {
        # ----------------------------------------------------
        # Replace HPModelTable in INI file with list from grid
        # ----------------------------------------------------
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
        CMTraceLog -Message " ... INI file not updated" -Type $TypeWarn
    } # else if ( Test-Path $lRepoLogFile )

    if ( $DebugMode ) { CMTraceLog -Message '< Update_INIModels' -Type $TypeNorm }
    
} # Function Update_INIModels 

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
    { ProdCode = '8723'; Model = 'HP ZBook Firefly 14 G7 Mobile Workstation' ; AddOns = 'Hotkeys','Notifications' }
#>
    $pModelsTable | 
        ForEach-Object {
            # populate checkmark, ProdId, Model Name, and then AddOns       
            [void]$pGrid.Rows.Add( @( $False, $_.ProdCode, $_.Model) )

            # also, add AddOns Platform flag file w/contents, if setting exists for the model

            if ( $_.AddOns ) { 
                $lRow = $pGrid.RowCount-1
                $pGrid.rows[$lRow].Cells['AddOns'].Value = $True
                Manage_IDFlag_File $pCurrentRepository $_.ProdCode $_.Model $pModelsTable[$lRow].AddOns $True
            } # if ( $_.AddOns )
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

    if ( $pCommonFlag ) {
        $Script:v_Root_CommonRepoFolder = $pNewPath
        $CommonRadioButton.Checked = $True
        $CommonPathTextField.Text = $pNewPath
        $CommonPathTextField.BackColor = $BackgroundColor
        $SharePathTextField.BackColor = ""
        $find = "^[\$]v_Root_CommonRepoFolder"
        $replace = "`$v_Root_CommonRepoFolder = ""$pNewPath"""
    } else {
        $Script:v_Root_IndividualRepoFolder = $pNewPath
        $SharedRadioButton.Checked = $True
        $SharePathTextField.Text = $pNewPath
        $SharePathTextField.BackColor = $BackgroundColor
        $CommonPathTextField.BackColor = ""
        $find = "^[\$]v_Root_IndividualRepoFolder"
        $replace = "`$v_Root_IndividualRepoFolder = ""$pNewPath"""
    } # else if ( $pCommonFlag )

    Mod_INISetting $IniFIleFullPath $find $replace "$($Script:IniFile) - New INI setting: ''$($replace)''"
    
    $Script:v_CommonRepo = $pCommonFlag

    # set $v_CommonRepo default in INI file
    $find = "^[\$]v_CommonRepo"
    $replace = "`$v_CommonRepo = `$$pCommonFlag"  # set up the replacing string to either $false or $true from ini file

    Mod_INISetting $IniFIleFullPath $find $replace "$($Script:IniFile) - New INI setting: ''$($replace)''"

} # Function Update_UIandINI

#=====================================================================================
<#
    Function Manage_IDFlag_File
    Parameters
        $pCreateFile                        # $True to create the flag file, $False to remove
        $pPath
        $pSysID
        $pModelName
#>
Function Manage_IDFlag_File {
    [CmdletBinding()]
	param( $pPath,
            $pSysID,                        # 4-digit hex-code platform/motherboard ID
            $pModelName,
            [array]$pAddOns,
            $pCreateFile)                    # $True = create file, $False = remove file
    if ( $Script:v_CommonRepo ) {
        $lFlagFile = $pPath+'\'+$AddSoftware+'\'+$pSysID
    } else {
        $lFlagFile = $pPath+'\'+$pModelName+'\'+$AddSoftware+'\'+$pSysID
    }
    if ( $pCreateFile ) {
        $lMsg = "... Setting download of AddOns Softpaqs for Platform: $pSysID"
        if ( -not (Test-Path $lFlagFile ) ) { 
            New-Item $lFlagFile 
            if ( $pAddOns.count -gt 0 ) { $pAddOns[0] | Out-File -FilePath $lFlagFile }
            For ($i=1; $i -le $pAddOns.count; $i++) { $pAddOns[$i] | Out-File -FilePath $lFlagFile -Append }
        } # if ( -not (Test-Path $lFlagFile ) )
    } else {
        $lMsg = "... Disabling download of AddOns Softpaqs for Platform: $pSysID"
        Remove-Item -Path $lFlagFile -ErrorAction silentlycontinue
    } # else if ( $pCreateFile )

    CMTraceLog -Message $lMsg -Type $TypeNorm

} # Manage_IDFlag_File

########################################################################################
#=====================================================================================
<#
    Function CreateForm
    This is the MAIN function with a Gui that sets things up for the user
#>
Function CreateForm {
    
    Add-Type -assembly System.Windows.Forms

    $LeftOffset = 20
    $TopOffset = 20
    $FieldHeight = 20
    $FormWidth = 900
    $FormHeight = 800

    $BackgroundColor = 'LightSteelBlue'
    
    #----------------------------------------------------------------------------------
    # ToolTips
    #----------------------------------------------------------------------------------
    
    $CM_tooltip = New-Object System.Windows.Forms.ToolTip
    $ShowHelp={
        #display popup help
        #each value is the name of a control on the form.
    
        Switch ($this.name) {
            "OS_Selection"    {$tip = "What Windows 10 OS version to work with"}
            "Keep Filters"    {$tip = "Do NOT erase previous product selection filters"}
            "Continue on 404" {$tip = "Continue Sync evern with Error 404, missing files"}
            "Individual Paths" {$tip = "Path to Head of Individual platform repositories"}
            "Common Path"     {$tip = "Path to Common/Shared platform repository"}
            "Models Table"    {$tip = "HP Models table to Sync repository(ies) to"}
            "Check All"       {$tip = "This check selects all Platforms and categories"}
            "Sync"            {$tip = "Syncronize repository for selected items from HP cloud"}
        } # Switch ($this.name)
        $CM_tooltip.SetToolTip($this,$tip)
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
    $PathsGroupBox.location = New-Object System.Drawing.Point(($LeftOffset+355),($TopOffset-10)) # (from left, from top)
    $PathsGroupBox.Size = New-Object System.Drawing.Size(($FormWidth-415),60)                    # (width, height)
    $PathsGroupBox.text = "Repository Paths - from $($IniFile):"
    
    #-------------------------------------------------------------------------------
    # create 'individual' radion button, label, text entry fields, and Browse button
    #-------------------------------------------------------------------------------

    $PathFieldLength = 300
    $labelWidth = 55

    if ( $DebugMode ) { Write-Host 'creating Shared radio button' }
    $SharedRadioButton = New-Object System.Windows.Forms.RadioButton
    $SharedRadioButton.Location = '10,14'
    $SharedRadioButton.Add_Click( {
            Import_IndividualRepos $dataGridView $SharePathTextField.Text
        }
    ) # $SharedRadioButton.Add_Click()

    if ( $DebugMode ) { Write-Host 'creating Individual field label' }
    $SharePathLabel = New-Object System.Windows.Forms.Label
    $SharePathLabel.Text = "Individual"
    #$SharePathLabel.TextAlign = "Left"    
    $SharePathLabel.location = New-Object System.Drawing.Point(($LeftOffset+15),$TopOffset) # (from left, from top)
    $SharePathLabel.Size = New-Object System.Drawing.Size($labelWidth,20)                   # (width, height)
    if ( $DebugMode ) { Write-Host 'creating Shared repo text field' }
    $SharePathTextField = New-Object System.Windows.Forms.TextBox
    $SharePathTextField.Text = "$Script:v_Root_IndividualRepoFolder"       # start w/INI setting
    $SharePathTextField.Multiline = $false 
    $SharePathTextField.location = New-Object System.Drawing.Point(($LeftOffset+80),($TopOffset-4)) # (from left, from top)
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
        Browse_IndividualRepos $dataGridView $SharePathTextField.Text
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
    $CommonPathTextField.location = New-Object System.Drawing.Point(($LeftOffset+80),($TopOffset+15)) # (from left, from top)
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
        [string]$lnewRepoFolder = Browse_CommonRepo $dataGridView $CommonPathTextField.Text
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
    # next 2 lines clear any selection from the initial data view
    #----------------------------------------------------------------------------------
    #$dataGridView.CurrentCell = $dataGridView[0,0]
    $dataGridView.ClearSelection()
    
    #----------------------------------------------------------------------------------
    # handle ALL checkbox selections here
    # columns= 0(row selected), 1(ProdCode), 2(Prod Name), 3-6(categories, 7(AddOns), 8(repository path)
    #----------------------------------------------------------------------------------
    $CellOnClick = {

        $row = $this.currentRow.index 
        $column = $this.currentCell.ColumnIndex 
        # checkmark cell returns a 'bool' $true or $false, so review those clicks
        $CurrCellValue = $dataGridView.rows[$row].Cells[$column].value
        
        # next, see if the cell is a checkmark (type Boolean) or a text cell (which would NOT have a value of $true or $false
        # columns 1=sysId, 2=Name, 8=path (all string types)
        if ( $column -in @(0, 3, 4, 5, 6, 7) ) {
        #if ( $dataGridView.rows[$row].Cells[$column].Value.GetType() -eq [Boolean] ) {

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
                        if ( $v_CommonRepo ) {
                            $lRepoPath = $CommonPathTextField.Text
                        } else {
                            $lRepoPath = $SharePathTextField.Text
                        }                      
                        Manage_IDFlag_File $lRepoPath $datagridview[1,$row].value $datagridview[2,$row].value $HPModelsTable[$row].AddOns $CellNewState 

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
    # Add a Refresh Grid from INI button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating a Get Models from INI button' }
    $GetModelsFromINIButton = New-Object System.Windows.Forms.Button
    $GetModelsFromINIButton.Width = 80
    $GetModelsFromINIButton.Height = 35
    $GetModelsFromINIButton.Location = New-Object System.Drawing.Point(($LeftOffset),($FormHeight-447))    # (from left, from top)
    $GetModelsFromINIButton.Text = 'Get INI Models'

    $GetModelsFromINIButton.add_Click( { 
        Empty_Grid $dataGridView
        Populate_Grid_from_INI $dataGridView $CommonPathTextField.Text $Script:HPModelsTable
    } )

    $CM_form.Controls.Add($GetModelsFromINIButton)

    #----------------------------------------------------------------------------------
    # Add a Refresh Grid button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating a Refresh Grid button' }
    $RefreshGridButton = New-Object System.Windows.Forms.Button
    $RefreshGridButton.Location = New-Object System.Drawing.Point(($LeftOffset+90),($FormHeight-447))    # (from left, from top)
    $RefreshGridButton.Text = 'Refresh Grid'
    $RefreshGridButton.AutoSize=$true

    $RefreshGridButton_Click={
        Get_Filters  $dataGridView $True
    } # $RefreshGridButton_Click={

    $RefreshGridButton.add_Click($RefreshGridButton_Click)

    $CM_form.Controls.Add($RefreshGridButton)
  
    #----------------------------------------------------------------------------------
    # Add a list filters button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating a list filters button' }
    $ListFiltersdButton = New-Object System.Windows.Forms.Button
    $ListFiltersdButton.Location = New-Object System.Drawing.Point(($LeftOffset+180),($FormHeight-447))    # (from left, from top)
    $ListFiltersdButton.Text = 'List Filters'
    $ListFiltersdButton.AutoSize=$true

    $ListFiltersdButton_Click={
        CMTraceLog -Message 'HPIA Repository Filters found...' -Type $TypeNorm
        Get_Filters  $dataGridView $False      # $False = do not refresh grid, just list the filters
    } # $ListFiltersdButton_Click={

    $ListFiltersdButton.add_Click($ListFiltersdButton_Click)

    $CM_form.Controls.Add($ListFiltersdButton)

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

        # Modify INI file with newly selected v_OSVER

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
                Backup_Log $LogFile
            }
            if ( $updateCMCheckbox.checked ) {
                if ( ($Script:CMConnected = Test_CMConnection) ) {
                    if ( $DebugMode ) { CMTraceLog -Message 'Script connected to CM' -Type $TypeDebug }
                }
            } # if ( $updateCMCheckbox.checked )
            if ( $Script:v_CommonRepo ) {
                Sync_Common_Repository $dataGridView $CommonPathTextField.Text $lCheckedListArray
            } else {
                sync_individual_repositories $dataGridView $lCheckedListArray
            }
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
    $updateCMCheckbox.Location = New-Object System.Drawing.Point(($LeftOffset),($TopOffset-5))   # (from left, from top)

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
    $CMDistributeheckbox.Location = New-Object System.Drawing.Point(($LeftOffset+120),($TopOffset-5))

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
    $HPIAPathField.location = New-Object System.Drawing.Point(($LeftOffset+100),($TopOffset-5)) # (from left, from top)
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
                    $CommonRadioButton.Checked = $true                          # set the visual default from the INI setting
                    $CommonPathTextField.BackColor = $BackgroundColor
                } else { 
                    $lIndividualRoot = $SharePathTextField.Text
                    if ( ([string]::IsNullOrEmpty($lIndividualRoot)) -or `
                         (-not (Test-Path $lIndividualRoot)) ) {
                        Write-Host "Individual Repository Path from INI file not found: $($lIndividualRoot)" -ForegroundColor Red
                        Init_Repository $lIndividualRoot $False            # False = just create the folder, will be a root of repos
                    }
                    Import_IndividualRepos $dataGridView $lIndividualRoot
                    $SharedRadioButton.Checked = $true 
                    $SharePathTextField.BackColor = $BackgroundColor
                } # else if ( $_ -match '\$true' )
            } # if ($_ -match $find)
        } # Foreach-Object

    #----------------------------------------------------------------------------------
    # Finally, show the dialog on screen
    #----------------------------------------------------------------------------------

    if ( $DebugMode ) { Write-Host 'calling ShowDialog' }
    $CM_form.ShowDialog() | Out-Null

} # Function CreateForm 

########################################################################################
# --------------------------------------------------------------------------
# Start of Invocation
# --------------------------------------------------------------------------

# at this point, we are past the -h | -help runstring option... 
# ... if any runstring options left, then we need to run without the UI

# in case we need to browse for a file, create the object now
	
#Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }

#$null = $FileBrowser.ShowDialog()

if ( $MyInvocation.BoundParameters.count -eq 0) {
    CreateForm                            # Create the GUI and take over all actions, like Report and Download
} else {
    $RunUI = $false

    if ( $PSBoundParameters.Keys.Contains('newLog') ) { Backup_Log $LogFile } 
    CMTraceLog -Message 'HPIARepo_Downloader - BEGIN'
    CMTraceLog -Message 'Script Path: '$ScriptPath
    CMTraceLog -Message 'Script Name: '$scriptName
    CMTraceLog -Message "$MyInvocation.BoundParameters.Keys: $($MyInvocation.BoundParameters.Keys)"
    CMTraceLog -Message "$MyInvocation.BoundParameters.Values: $($MyInvocation.BoundParameters.Values)"
    if ( $PSBoundParameters.Keys.Contains('IniFile') ) { '-iniFile: ' + $inifile }
    if ( $PSBoundParameters.Keys.Contains('RepoStyle') ) { if ( $RepoStyle -match 'Common' ) { '-RepoStyle: ' + $RepoStyle
            $v_CommonRepo = $true } else { $v_CommonRepo = $false }  } 
    if ( $PSBoundParameters.Keys.Contains('Products') ) { "-Products: $($Products)" }
    if ( $PSBoundParameters.Keys.Contains('ListFilters') ) { list_filters }
    if ( $PSBoundParameters.Keys.Contains('NoIniSw') ) { '-NoIniSw' }
    if ( $PSBoundParameters.Keys.Contains('showActivityLog') ) { $showActivityLog = $true } 
    if ( $PSBoundParameters.Keys.Contains('Sync') ) { sync_repos }

} # if ( $MyInvocation.BoundParameters.count -gt 0)

########################################################################################