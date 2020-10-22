<#
    HP Image Assistant and Softpaq Repository Downloader
    by Dan Felman/HP Technical Consultant
        
        Loop Code: The HPModelsTable loop code, now in separate INI.ps1 file (and other general code) was taken from Gary Blok's (@gwblok) post on garytown.com.
        https://garytown.com/create-hp-bios-repository-using-powershell
 
        Logging: The Log function was created by Ryan Ephgrave (@ephingposh)
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
        Added 'Distribute SCCM Packages' ($DistributeCMPackages) variable use in INI file
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
        Added 'INI S/W' column to allow/disallow software listed in INI.ps1 file from being downloaded
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
#>
<#
    Args
        -inifile - full path to ini.ps1 file , also 1st arg in command line
#>
[CmdletBinding()]
param(
    [Parameter( Mandatory = $false, Position = 0, ValueFromRemainingArguments )]
    [Switch]$Help,                                # $help is set to $true if '-help' passed as argument

    [Parameter( Mandatory = $false, Position = 1, HelpMessage="Path to ini.ps1 file. Default is 'HPIARepo_ini.ps1'" )]
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

$ScriptVersion = "1.70 (10/20/2020)"

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
'invoked from: '+$ScriptPath

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
. $IniFIleFullPath                                   # source the code in the INI file      

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
        $LogMessage | Out-File -Append -Encoding UTF8 -FilePath $Script:LogFile

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

    CMTraceLog -Message "Checking for HP CMSL... " -Type $TypeNoNewline

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
                    $CMGroupBox.Text = 'SCCM - Connected'
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
        #if ( $DebugMode ) { CMTraceLog -Message "... setting HPIA Package to Version: $($OSVER), path: $($pRepoPath)" -Type $TypeDebug }

        CMTraceLog -Message "... HPIA package - setting name ''$pHPIAPkgName'', version ''$pHPIAVersion'', Path ''$pHPIAPath''" -Type $TypeNorm
	    Set-CMPackage -Name $pHPIAPkgName -Version $pHPIAVersion
	    Set-CMPackage -Name $pHPIAPkgName -Path $pHPIAPath

        if ( $Script:DistributeCMPackages  ) {
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
    if ( $DebugMode ) { CMTraceLog -Message "... setting CM Package to Version: $($OSVER), path: $($pRepoPath)" -Type $TypeDebug }

	Set-CMPackage -Name $lPkgName -Version "$($OSVER)"
	Set-CMPackage -Name $lPkgName -Path $pRepoPath

    if ( $Script:DistributeCMPackages  ) {
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
            }
        } # for ($i=0; $i -lt 100; $i++ )
    } # if ( $pLogFileFullPath )

} # Backup_Log

#=====================================================================================
<#
    Function init_repository
        This function will create a repository folder and initialize it for HPIA, if necessary
        Args:
            $pRepoFOlder: folder to validate, or create
            $pInitRepository: variable - if $true, initialize repository, otherwise, ignore
#>
Function init_repository {
    [CmdletBinding()]
	param( $pRepoFolder,
            $pInitRepository )

    if ( $DebugMode ) { CMTraceLog -Message "> init_repository" -Type $TypeNorm }

    $lCurrentLoc = Get-Location

    $retRepoCreated = $false

    if ( (Test-Path $pRepoFolder) -and ($pInitRepository -eq $false) ) {
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
        } # if ( !(Test-Path $RepoShareMain) )

        $retRepoCreated = $true

        #--------------------------------------------------------------------------------
        # if needed, check on repository to initialize (CMSL repositories have a .Repository folder)

        if ( $pInitRepository -and !(test-path "$pRepoFolder\.Repository")) {
            Set-Location $pRepoFolder
            $initOut = (Initialize-Repository) 6>&1
            CMTraceLog -Message  "... Repository Initialization done: $($Initout)"  -Type $TypeNorm 

            CMTraceLog -Message  '... configuring this repository for HP Image Assistant' -Type $TypeNorm
            Set-RepositoryConfiguration -setting OfflineCacheMode -cachevalue Enable 6>&1   # configuring the repo for HP IA's use

        } # if ( $pInitRepository -and !(test-path "$pRepoFOlder\.Repository"))

    } # else if ( (Test-Path $pRepoFolder) -and ($pInitRepository -eq $false) )

    #--------------------------------------------------------------------------------
    # intialize special folder for holding named softpaqs
    # ... this folder will hold softpaqs added outside CMSL sync/cleanup folder
    #--------------------------------------------------------------------------------
    $lAddSoftpaqsFolder = "$($pRepoFolder)\$($AddSoftware)"

    if ( !(Test-Path $lAddSoftpaqsFolder) -and $pInitRepository ) {
        CMTraceLog -Message "... creating Addon Softpaq Folder $lAddSoftpaqsFolder" -Type $TypeNorm
        New-Item -Path $lAddSoftpaqsFolder -ItemType directory
        if ( $DebugMode ) { CMTraceLog -Message "NO $lAddSoftpaqsFolder" -Type $TypeWarn }
    } # if ( !(Test-Path $lAddSoftpaqsFolder) )

    #Set-Location $lCurrentLoc

    if ( $DebugMode ) { CMTraceLog -Message "< init_repository" -Type $TypeNorm }
    return $retRepoCreated

} # Function init_repository

#=====================================================================================
<#
    Function Get_HPServerIP
        This function will run a Sync and a Cleanup commands from HPCMSL

    expects parameter 
        - string of function calling
        - file to store info for further analysis
#>
Function Get_HPServerIP {
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
        CMTraceLog -Message '[Get_HPServerIP] proxy IS enabled - Can''t obtain real HP Server IP' -Type $TypeWarn 
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

} # Get_HPServerIP

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

    if ( $script:noIniSw ) {
        return
    }
    if ( $DebugMode ) { CMTraceLog -Message  '> download_softpaqs_by_name' }

    # let's search the HPModels list from the INI.ps1 file for Software entries

    $HPModelsTable  | 
        ForEach-Object {
            if ( ($_.ProdCode -match $pProdCode) -and $pINISWSelected ) {                      # let's match the Model to download for
                
                $lAddSoftpaqsFolder = "$($pFolder)\$($AddSoftware)"
                Set-Location $lAddSoftpaqsFolder
                CMTraceLog -Message "... Retrieving named Softpaqs for ProdCode:$pProdCode" -Type $TypeNorm

                ForEach ( $Softpaq in $_.SqName ) {

                    Try {
                        Get_HPServerIP 'get-softpaqlist' $HPServerIPFile 2      # 2 secs delay
                        $lListret = (Get-SoftpaqList -platform $pProdCode -os $OS -osver $OSVER -characteristic SSM) 6>&1 | 
                            Where-Object { ($_.Name -match $Softpaq) -or ($_.Id -match $Softpaq) } | 
                            ForEach-Object { 
                                if ($_.SSM -eq $true) {
                                    #CMTraceLog -Message "      [Get Softpaq] - $($_.Id) ''$Softpaq''" -Type $TypeNoNewLine
                                    CMTraceLog -Message "      [Get Softpaq] - $($_.Id) ''$($_.Name)''" -Type $TypeNoNewLine
                                    if ( Test-Path "$($_.Id).exe" ) {
                                        CMTraceLog -Message  " - Already exists. Will not re-download"  -Type $TypeWarn
                                    } else {
                                        
                                        Get_HPServerIP 'get-softpaq' $HPServerIPFile 2
                                        $ret = (Get-Softpaq $_.Id) 6>&1
                                        CMTraceLog -Message  "- Downloaded $($ret)"  -Type $TypeWarn 
                                        #--------------------------------------------------------------------------------
                                        # see if we captured a connection to HP                               
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
                                } # if ($_.SSM -eq $true)

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
                        $lerrMsg = "      [Get-SoftpaqList] $($lListret)... 'exception -platform:$($pProdCode) -osver:$($OSVER) - Filter NOT Supported!" 
                        CMTraceLog -Message $lerrMsg -Type $TypeError
                    }
                    
                } # ForEach ( $Softpaq in $_.SqName )

            } # if ( $_.ProdCode -match $pProdCode )
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
        Get_HPServerIP 'invoke-repositorysync' $HPServerIPFile 2
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
            CMTraceLog -Message "... Copying named softpaqs to Repository - if selected"
            $lRobocopySource = "$($pFolder)\$($AddSoftware)"
            $lRobocopyDest = "$($pFolder)"
            $lRobocopyArg = '"'+$lRobocopySource+'"'+' "'+$lRobocopyDest+'"'
            $RobocopyCmd = "robocopy.exe"
            Start-Process -FilePath $RobocopyCmd -ArgumentList $lRobocopyArg -WindowStyle Hidden -Wait     
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
#>

#Update_Repository_and_Grid $lMainRepo $lModelId $item $true    # $true means 'add filters'

Function Update_Repository_and_Grid {
[CmdletBinding()]
	param( $pModelRepository,
            $pModelID,
            $pRow,
            $pAddFilters )

    $pCurrentLoc = Get-Location

    # move to location of Repository to use CMSL repo commands
    set-location $pModelRepository

    if ( $Script:Continueon404 ) {
        Set-RepositoryConfiguration -Setting OnRemoteFileNotFound -Value LogAndContinue
    } else {
        Set-RepositoryConfiguration -Setting OnRemoteFileNotFound -Value Fail
    }

    if ( $pAddFilters ) {
    
        if ( $Script:KeepFilters ) {
            CMTraceLog -Message  "... Keeping existing filters in: $($pModelRepository) -platform $($pModelID)" -Type $TypeNorm
        } else {
            CMTraceLog -Message  "... Removing existing filters in: $($pModelRepository) -platform $($pModelID)" -Type $TypeNorm
            $lres = (Remove-RepositoryFilter -platform $pModelID -yes 6>&1)      
            if ( $debugMode ) { CMTraceLog -Message "... removed filter: $($lres)" -Type $TypeWarn }
        } # else if ( $Script:KeepFilters )
        #--------------------------------------------------------------------------------
        # now update filters - every category checked for the current model'
        #--------------------------------------------------------------------------------

        if ( $DebugMode ) { CMTraceLog -Message "... adding category filters" -Type $TypeDebug }

        foreach ( $cat in $FilterCategories ) {

            if ( $datagridview.Rows[$pRow].Cells[$cat].Value ) {
                CMTraceLog -Message  "... adding filter: -platform $($pModelID) -os $OS -osver $OSVER -category $($cat) -characteristic ssm" -Type $TypeNoNewline
                $lRes = (Add-RepositoryFilter -platform $pModelID -os win10 -osver $OSVER -category $cat -characteristic ssm 6>&1)
                CMTraceLog -Message $lRes -Type $TypeWarn 
            }
        } # foreach ( $cat in $FilterCategories )

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

        foreach ( $filterSetting in $lProdFilters ) {
            <#
                platform        : 83B2
                operatingSystem : win10:1909
                category        : Firmware
                releaseType     : *
                characteristic  : ssm
            #>
            if ( $filterSetting.platform -match $pModelID ) {
                CMTraceLog -Message "... unselected platform $($filterSetting.platform) matched - filters removed" -Type $TypeWarn 
                $lres = (Remove-RepositoryFilter -platform $pModelID -yes 6>&1)      
                if ( $debugMode ) { CMTraceLog -Message "... removed filter: $($lres)" -Type $TypeWarn }
            }

        } # foreach ( $filterSetting in $lProdFilters )

    } # else if ( $pAddFilters ) {

    Set-Location -Path $pCurrentLoc

} # Update_Repository_and_Grid

#=====================================================================================
<#
    Function Sync_Repos
    
#>
Function Sync_Repos {

    $lCurrentSetLoc = Get-Location

    #--------------------------------------------------------------------------------
    # find out if the share exists, if not, just return
    #--------------------------------------------------------------------------------

    if ( $CommonRepo ) {
            
        # basic check to confirm the repository was configured for HPIA
        if ( !(Test-Path "$($RepoShareCommon)\.Repository") ) {
            CMTraceLog -Message  "... Common Repository Folder selected, not initialized" -Type $TypeNorm
            return
        } 
        set-location $RepoShareCommon                 
        Sync_and_Cleanup_Repository $RepoShareCommon  

    } else {

        # basic check to confirm the repository exists that hosts individual repos
        if ( Test-Path $RepoShareMain ) {
            # let's traverse every product Repository folder
            $lProdFolders = Get-ChildItem -Path $RepoShareMain | where {($_.psiscontainer)}

            foreach ( $lprodName in $lProdFolders ) {
                $lCurrentPath = "$($RepoShareMain)\$($lprodName.name)"
                set-location $lCurrentPath
                Sync_and_Cleanup_Repository $lCurrentPath  
            } # foreach ( $lprodName in $lProdFolders )
        } else {
            CMTraceLog -Message  "... Shared/Individual Repository Folder selected, Head repository not initialized" -Type $TypeNorm
        } # else if ( !(Test-Path $RepoShareMain) ) 
    } # else if ( $Script:CommonRepo )

    CMTraceLog -Message  "Sync DONE" -Type $TypeSuccess
    Set-Location -Path $lCurrentSetLoc

} # Sync_Repos

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
	param( $pModelsList,                                             # array of row lines that are checked
            $pCheckedItemsList,                                      # array of rows selected
            $pRemoveFilters,
            $pContinueOnError)
    
    $pCurrentLoc = Get-Location
    CMTraceLog -Message "> Sync_Common_Repository - START" -Type $TypeNorm
    if ( $DebugMode ) { CMTraceLog -Message "... for OS version: $($OSVER)" -Type $TypeDebug }

    $lMainRepo =  $RepoShareCommon   
    init_repository $lMainRepo $true                                 # make sure Main repo folder exists, or create it - no init
    CMTraceLog -Message  "... Common repository selected: $($lMainRepo)" -Type $TypeNorm

    #--------------------------------------------------------------------------------
    if ( $DebugMode ) { CMTraceLog -Message "... stepping through selected models" -Type $TypeDebug }

    # loop through every selected Model in the list

    for ( $item = 0; $item -lt $pModelsList.RowCount; $item++ ) {
        
        $lModelId = $pModelsList[1,$item].Value                            # column 1 has the Model/Prod ID
        $lModelName = $pModelsList[2,$item].Value                          # column 2 has the Model name

        # if model entry row is checked, we need to create a repository, but ONLY if $CommonRepo = $false

        if ( $item -in $pCheckedItemsList ) {

            CMTraceLog -Message "--- Updating model: $lModelName"

            # update repo filters and show in grid
            Update_Repository_and_Grid $lMainRepo $lModelId $item $true    # $true means 'add filters'
            #--------------------------------------------------------------------------------
            # update SCCM Repository package, if user decided with the UI checkmark
            #--------------------------------------------------------------------------------
            if ( $Script:UpdateCMPackages ) {
                CM_RepoUpdate $lModelName $lModelId $lMainRepo
            }
            #--------------------------------------------------------------------------------
            # update specific Softpaqs as defined in INI file for the current model
            #--------------------------------------------------------------------------------
            download_softpaqs_by_name  $lMainRepo $lModelID $pModelsList.Rows[$item].Cells['INI Software'].Value

        } else {

            # the product/model was not selected at this point
            if ( $pRemoveFilters ) {
                # make sure previous filters from 'unselected' models are removed from repository
                Update_Repository_and_Grid $lMainRepo $lModelId $item $false   # $false means 'do not add filters, just remove them'
            }

        } # if ( $item -in $pCheckedItemsList )

    } # for ( $item = 0; $item -lt $pModelsList.RowCount; $item++ )

    # we are done checking every model for filters, so cleanup

    Sync_and_Cleanup_Repository $lMainRepo

    #--------------------------------------------------------------------------------
    Set-Location -Path $pCurrentLoc

    CMTraceLog -Message "< Sync_Common_Repository DONE" -Type $TypeSuccess

} # Function Sync_Common_Repository

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
    if ( $DebugMode ) { CMTraceLog -Message "... for OS version: $($OSVER)" -Type $TypeDebug }

    #--------------------------------------------------------------------------------
    # decide if we work on single repo or individual repos
    #--------------------------------------------------------------------------------

    $lMainRepo =  $RepoShareMain                                 # 
    init_repository $lMainRepo $false                            # make sure Main repo folder exists, do NOT make it a repository

    #--------------------------------------------------------------------------------
    if ( $DebugMode ) { CMTraceLog -Message "... stepping through all selected models" -Type $TypeDebug }

    # go through every selected HP Model in the list

    for ( $i = 0; $i -lt $pModelsList.RowCount; $i++ ) {
        
        $lModelId = $pModelsList[1,$i].Value                      # column 1 has the Model/Prod ID
        $lModelName = $pModelsList[2,$i].Value                    # column 2 has the Model name

        # if model entry is checked, we need to create a repository, but ONLY if $CommonRepo = $false
        if ( $i -in $pCheckedItemsList ) {

            CMTraceLog -Message "`n--- Updating model: $lModelName"

            $lTempRepoFolder = "$($lMainRepo)\$($lModelName)"     # this is the repo folder for this model
            init_repository $lTempRepoFolder $true
            
            Update_Repository_and_Grid $lTempRepoFolder $lModelId $i $true    # $true means 'add filters'
            
            #--------------------------------------------------------------------------------
            # update specific Softpaqs as defined in INI file for the current model
            #--------------------------------------------------------------------------------
            download_softpaqs_by_name  $lTempRepoFolder $lModelID $pModelsList.Rows[$i].Cells['INI Software'].Value

            #--------------------------------------------------------------------------------
            # now sync up and cleanup this repository
            #--------------------------------------------------------------------------------
            Sync_and_Cleanup_Repository $lTempRepoFolder 

            #--------------------------------------------------------------------------------
            # update SCCM Repository package, if user allows
            #--------------------------------------------------------------------------------
            if ( $Script:UpdateCMPackages ) {
                CM_RepoUpdate $lModelName $lModelId $lTempRepoFolder
            }
            
        } # if ( $i -in $pCheckedItemsList )

    } # for ( $i = 0; $i -lt $pModelsList.RowCount; $i++ )

    CMTraceLog -Message "< sync_individual_repositories DONE" -Type $TypeSuccess

} # Function sync_individual_repositories

#=====================================================================================
<#
    Function clear_datagrid
        clear all checkmarks and last path column '8', except for SysID and Model columns

#>
Function clear_datagrid {
    [CmdletBinding()]
	param( $pDataGrid )                             

    if ( $DebugMode ) { CMTraceLog -Message '> clear_datagrid' -Type $TypeNorm }

    for ( $row = 0 ; $row -lt $pDataGrid.RowCount ; $row++ ) {
        for ( $col = 0 ; $col -lt $pDataGrid.ColumnCount ; $col++ ) {
            if ( $col -in @(0,3,4,5,6,7) ) {
                $pDataGrid[$col,$row].value = $false                       # clear checkmarks
            } else {
                if ( $pDataGrid.columns[$col].Name -match 'Repo' ) {
                    $pDataGrid[$col,$row].value = ''                       # clear path text field
                }
            } # else if ( $col -in @(0,3,4,5,6,7) )
        }
    } # for ( $row = 0 ; $row -lt $pDataGrid.RowCount ; $row++ ) 
    
    if ( $DebugMode ) { CMTraceLog -Message '< clear_datagrid' -Type $TypeNorm }

} # Function clear_datagrid

#=====================================================================================
<#
    Function list_filters
        List filters out to console... function called from runstring

#>
Function list_filters {

    $lCurrentSetLoc = Get-Location

    #--------------------------------------------------------------------------------
    # find out if the share exists, if not, just return
    #--------------------------------------------------------------------------------

    if ( $CommonRepo ) {
            
        # basic check to confirm the repository was configured for HPIA
        if ( !(Test-Path "$($RepoShareCommon)\.Repository") ) {
            CMTraceLog -Message "... Common Repository Folder selected, not initialized" -Type $TypeNorm 
            return
        } 
        set-location $RepoShareCommon
            
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

        # basic check to confirm the repository was configured for HPIA
        if ( !(Test-Path $RepoShareMain) ) {
            CMTraceLog -Message "... Shared/Individual Repository Folder selected, Head repository not initialized" -Type $TypeNorm
            "... Shared/Individual Repository Folder selected, Head repository not initialized"
            return
        } 
        set-location $RepoShareMain | where {($_.psiscontainer)}

        # let's traverse every product Repository folder
        $lProdFolders = Get-ChildItem -Path $RepoShareMain

        foreach ( $lprodName in $lProdFolders ) {
            set-location "$($RepoShareMain)\$($lprodName.name)"
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

    } # else if ( $Script:CommonRepo )

    Set-Location -Path $lCurrentSetLoc

} # list_filters

#=====================================================================================
<#
    Function get_filters
        Retrieves category filters from the share for each selected model
        ... and populates the Grid appropriately

#>
Function get_filters {
    [CmdletBinding()]
	param( $pDataGridList,                                  # array of row lines that are checked
        $pListFiltersOnly )                                 # if $true, list filters, don't refresh grid

    $pCurrentLoc = Get-Location

    if ( $DebugMode ) { CMTraceLog -Message '> get_filters' -Type $TypeNorm }

    if ( -not $pListFiltersOnly ) {
        clear_datagrid $pDataGridList
    }

    #--------------------------------------------------------------------------------
    # find out if the share exists, if not, just return
    #--------------------------------------------------------------------------------
    if ( (Test-Path $RepoShareMain) -or (Test-Path $RepoShareCommon) ) {

        #--------------------------------------------------------------------------------
        if ( $Script:CommonRepo ) {
            # see if the repository was configured for HPIA

            if ( !(Test-Path "$($RepoShareCommon)\.Repository") ) {
                CMTraceLog -Message "... Repository Folder not initialized" -Type $TypeWarn
                if ( $DebugMode ) { CMTraceLog -Message '< get_filters - Done' -Type $TypeNorm }
                return
            } 
            set-location $RepoShareCommon
            if ( $pListFiltersOnly ) {
                CMTraceLog -Message "... Filters from Common Repository ...$($RepoShareCommon)" -Type $TypeNorm
            } else {
                CMTraceLog -Message '... Refreshing Grid from Common Repository ...' -Type $TypeNorm
            }
            
            $lProdFilters = (get-repositoryinfo).Filters

            foreach ( $filterSetting in $lProdFilters ) {

                if ( $DebugMode ) { CMTraceLog -Message "... populating filter categories ''$($lModelName)''" -Type $TypeDebug }

                # check each row SysID against the Filter Platform ID
                for ( $i = 0; $i -lt $pDataGridList.RowCount; $i++ ) {

                    if ( $filterSetting.platform -eq $pDataGridList[1,$i].value) {
                        if ( $pListFiltersOnly ) {
                            $lMsg = "... Platform $($filterSetting.platform) ... $($filterSetting.operatingSystem) $($filterSetting.category) $($filterSetting.characteristic)"
                            CMTraceLog -Message $lMsg -Type $TypeWarn
                        } else {
                            # we matched the row/SysId with the Filter Platform IF, so let's add each category in the filter
                            foreach ( $cat in  ($filterSetting.category.split(' ')) ) {
                                $pDataGridList.Rows[$i].Cells[$cat].Value = $true
                            }
                            $pDataGridList[0,$i].Value = $true
                            $pDataGridList[($dataGridView.ColumnCount-1),$i].Value = $RepoShareCommon
                        } # else if ( $pListFilters )
                    } # if ( $filterSetting.platform -eq $pDataGridList[1,$i].value)

                } # for ( $i = 0; $i -lt $pDataGridList.RowCount; $i++ )
            } # foreach ( $platform in $lProdFilters )

        } else {
            if ( -not $pListFiltersOnly ) {
                CMTraceLog -Message '... Refreshing Grid from Individual Repositories ...' -Type $TypeNorm
            }
            #--------------------------------------------------------------------------------
            # now check for each product's repository folder
            # if the repo is created, then check the category filters
            #--------------------------------------------------------------------------------
            for ( $i = 0; $i -lt $pDataGridList.RowCount; $i++ ) {

                $lModelId = $pDataGridList[1,$i].Value                                                # column 1 has the Model/Prod ID
                $lModelName = $pDataGridList[2,$i].Value                                              # column 2 has the Model name

                $lTempRepoFolder = "$($RepoShareMain )\$($lModelName)"                               # this is the repo folder for this model

                # move to location of Repository to use CMSL repo commands
                if ( Test-Path $lTempRepoFolder ) {
                    set-location $lTempRepoFolder

                    ###### filters sample obtained with get-repositoryinfo: 
                    ###### platform        : 8438
                    ###### operatingSystem : win10:2004 win10:2004
                    ###### category        : BIOS firmware
                    ###### releaseType     : *
                    ###### characteristic  : ssm

                    $lProdFilters = (get-repositoryinfo).Filters

                    foreach ( $platform in $lProdFilters ) {
                        if ( $pListFiltersOnly ) {
                            CMTraceLog -Message "... Platform $($lProdFilters.platform) ... $($lProdFilters.operatingSystem) $($lProdFilters.category) $($lProdFilters.characteristic) - @$($lTempRepoFolder)" -Type $TypeWarn
                        } else {
                            CMTraceLog -Message "... populating filter categories ''$($lModelName)''" -Type $TypeDebug
                        
                            foreach ( $cat in  ($lProdFilters.category.split(' ')) ) {
                                $pDataGridList.Rows[$i].Cells[$cat].Value = $true
                            }
                            #--------------------------------------------------------------------------------
                            # show repository path for this model (model is checked) (col 8 is links col)
                            #--------------------------------------------------------------------------------
                            $pDataGridList[0,$i].Value = $true
                            $pDataGridList[($dataGridView.ColumnCount-1),$i].Value = $lTempRepoFolder
                        }
                    } # foreach ( $platform in $lProdFilters )

                } # if ( Test-Path $lTempRepoFolder )
            } # for ( $i = 0; $i -lt $pModelsList.RowCount; $i++ ) 


        } # if ( $Script:CommonRepo )

    } else {
        if ( $DebugMode ) { CMTraceLog -Message 'Main Repo Host Folder ''$($lMainRepo)'' does NOT exist' -Type $TypeDebug }
    } # else if ( !(Test-Path $lMainRepo) )

    if ( $DebugMode ) { CMTraceLog -Message '< get_filters]' -Type $TypeNorm }
    
} # Function get_filters

#=====================================================================================
<#
    Function Mod_INISetting
    code to modify a line in the INI.ps1 file
    search for a setting name (find), change it to a new setting (replace)
#>
Function Mod_INISetting {
    [CmdletBinding()]
	param( $pINIFile, $pfind, $pReplace, $pText )

    CMTraceLog -Message $pText -Type $TypeNorm
    
    (Get-Content $pINIFile) | Foreach-Object {if ($_ -match $find) {$replace} else {$_}} | Set-Content $pINIFile

} # Mod_INISetting


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
    if ( $DebugMode ) { Write-Host 'creating Form' }
    $CM_form = New-Object System.Windows.Forms.Form
    $CM_form.Text = "HPIARepo_Downloader v$($ScriptVersion)"
    $CM_form.Width = $FormWidth
    $CM_form.height = $FormHeight
    $CM_form.Autosize = $true
    $CM_form.StartPosition = 'CenterScreen'

    #----------------------------------------------------------------------------------
    # Create Sync button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Sync button' }
    $buttonSync = New-Object System.Windows.Forms.Button
    $buttonSync.Width = 60
    $buttonSync.Text = 'Sync'
    $buttonSync.Location = New-Object System.Drawing.Point($LeftOffset,($TopOffset-1))

    $buttonSync.add_click( {

        # Modify INI file with newly selected OSVER

        if ( $Script:OSVER -ne $OSVERComboBox.Text ) {
            $find = "^[\$]OSVER"
            $replace = "`$OSVER = ""$($OSVERComboBox.Text)"""  
            Mod_INISetting $IniFIleFullPath $find $replace 'Changing INI file with selected OSVER'
            #(Get-Content $IniFIleFullPath) | Foreach-Object {if ($_ -match $find) {$replace} else {$_}} | Set-Content $IniFIleFullPath
        } 

        $Script:OSVER = $OSVERComboBox.Text                            # get selected version
        if ( $Script:OSVER -in $Script:OSVALID ) {
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
            if ( $Script:CommonRepo ) {
                $lRemoveFilters = $false
                $lResponse = [System.Windows.MessageBox]::Show("Remove existing filters for products not selected?","Common Repository",4)    # 1 = "OKCancel" ; 4 = "YesNo"
                if ( $lResponse -eq 'Yes' ) {
                    $lRemoveFilters = $true
                } 
                Sync_Common_Repository $dataGridView $lCheckedListArray $lRemoveFilters
            } else {
                sync_individual_repositories $dataGridView $lCheckedListArray
            }
        } # if ( $Script:OSVER -in $Script:OSVALID )

    } ) # $buttonSync.add_click

    #$CM_form.Controls.AddRange(@($buttonSync, $ActionComboBox))
    $CM_form.Controls.AddRange(@($buttonSync))

    #----------------------------------------------------------------------------------
    # Create OS and OS Version display fields - info from .ini file
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating OS Combo and Label' }
    $OSTextLabel = New-Object System.Windows.Forms.Label
    $OSTextLabel.Text = "Windows 10:"
    $OSTextLabel.location = New-Object System.Drawing.Point(($LeftOffset+70),($TopOffset+4))    # (from left, from top)
    $OSTextLabel.AutoSize = $true
    $OSVERComboBox = New-Object System.Windows.Forms.ComboBox
    $OSVERComboBox.Size = New-Object System.Drawing.Size(60,$FieldHeight)                  # (width, height)
    $OSVERComboBox.Location  = New-Object System.Drawing.Point(($LeftOffset+140), ($TopOffset))
    $OSVERComboBox.DropDownStyle = "DropDownList"
    $OSVERComboBox.Name = "OS_Selection"
    $OSVERComboBox.add_MouseHover($ShowHelp)
    # populate menu list from INI file
    Foreach ($MenuItem in $OSVALID) {
        [void]$OSVERComboBox.Items.Add($MenuItem);
    }  
    $OSVERComboBox.SelectedItem = $OSVER 

    $CM_form.Controls.AddRange(@($OSTextLabel,$OSVERComboBox))

    #----------------------------------------------------------------------------------
    # Create Keep Filters checkbox
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Keep Filters Checkbox' }
    $keepFiltersCheckbox = New-Object System.Windows.Forms.CheckBox
    $keepFiltersCheckbox.Text = 'Keep prev OS Filters'
    $keepFiltersCheckbox.Autosize = $true
    $keepFiltersCheckbox.Location = New-Object System.Drawing.Point(($LeftOffset+210),($TopOffset+3))   # (from left, from top)

    # populate CM Udate checkbox from .INI variable setting - $UpdateCMPackages
    $find = "^[\$]KeepFilters"
    (Get-Content $IniFIleFullPath) | Foreach-Object { 
        if ($_ -match $find) { 
            if ( $_ -match '\$true' ) { 
                $keepFiltersCheckbox.Checked = $true 
            } else { 
                $keepFiltersCheckbox.Checked = $false 
            }              } 
        } # Foreach-Object

    $Script:KeepFilters = $keepFiltersCheckbox.Checked

    # update .INI variable setting - $UpdateCMPackages 
    $keepFiltersCheckbox_Click = {
        $find = "^[\$]KeepFilters"
        if ( $keepFiltersCheckbox.checked ) {
            $Script:KeepFilters = $true
            CMTraceLog -Message "... Existing OS Filters will NOT be removed'" -Type $TypeWarn
        } else {
            $Script:KeepFilters = $false
            $keepFiltersCheckbox.Checked = $false
            CMTraceLog -Message "... Existing Filters will be removed and new filters will be created'" -Type $TypeWarn
        }
        $replace = "`$KeepFilters = `$$Script:KeepFilters"                   # set up the replacing string to either $false or $true from ini file
        Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"

    } # $keepFiltersCheckbox_Click = 

    $keepFiltersCheckbox.add_Click($keepFiltersCheckbox_Click)

    $CM_form.Controls.AddRange(@($keepFiltersCheckbox))

    #----------------------------------------------------------------------------------
    # Create Continue on Error checkbox
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Keep Filters Checkbox' }
    $continueOn404Checkbox = New-Object System.Windows.Forms.CheckBox
    $continueOn404Checkbox.Text = 'Continue on Sync Error'
    $continueOn404Checkbox.Autosize = $true
    $continueOn404Checkbox.Location = New-Object System.Drawing.Point(($LeftOffset+210),($TopOffset+23))   # (from left, from top)

    # populate CM Udate checkbox from .INI variable setting - $UpdateCMPackages
    $find = "^[\$]Continueon404"
    (Get-Content $IniFIleFullPath) | Foreach-Object { 
        if ($_ -match $find) { 
            if ( $_ -match '\$true' ) { 
                $continueOn404Checkbox.Checked = $true 
            } else { 
                $continueOn404Checkbox.Checked = $false 
            }              } 
        } # Foreach-Object

    $Script:Continueon404 = $continueOn404Checkbox.Checked

    # update .INI variable setting - $UpdateCMPackages 
    $continueOn404Checkbox_Click = {
        $find = "^[\$]Continueon404"
        if ( $continueOn404Checkbox.checked ) {
            $Script:Continueon404 = $true
            CMTraceLog -Message "... Will continue on Sync missing file errors'" -Type $TypeWarn
        } else {
            $Script:Continueon404 = $false
            $continueOn404Checkbox.Checked = $false
            CMTraceLog -Message "... Will STOP on Sync missing file errors'" -Type $TypeWarn
        }
        $replace = "`$Continueon404 = `$$Script:Continueon404"                   # set up the replacing string to either $false or $true from ini file
        Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"

    } # $continueOn404Checkbox_Click = 

    $continueOn404Checkbox.add_Click($continueOn404Checkbox_Click)

    $CM_form.Controls.AddRange(@($continueOn404Checkbox))

    #----------------------------------------------------------------------------------
    # add share and Common info fields and Radio Buttons for selection
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Share field' }
    $SharePathLabel = New-Object System.Windows.Forms.Label
    $SharePathLabel.Text = "Individual"
    $SharePathLabel.location = New-Object System.Drawing.Point(($LeftOffset+15),$TopOffset) # (from left, from top)
    $SharePathLabel.Size = New-Object System.Drawing.Size(55,20)                            # (width, height)
    #$SharePathLabel.TextAlign = "Left"
    $SharePathTextField = New-Object System.Windows.Forms.TextBox
    $SharePathTextField.Text = "$RepoShareMain"
    $SharePathTextField.Multiline = $false 
    $SharePathTextField.location = New-Object System.Drawing.Point(($LeftOffset+80),($TopOffset-4)) # (from left, from top)
    $SharePathTextField.Size = New-Object System.Drawing.Size(380,$FieldHeight)             # (width, height)
    $SharePathTextField.ReadOnly = $true
    $SharePathTextField.Name = "Share_Path"

    $SharePathLabelSingle = New-Object System.Windows.Forms.Label
    $SharePathLabelSingle.Text = "Common"
    $SharePathLabelSingle.location = New-Object System.Drawing.Point(($LeftOffset+15),($TopOffset+18)) # (from left, from top)
    $SharePathLabelSingle.Size = New-Object System.Drawing.Size(55,20)                            # (width, height)

    $SharePathSingleTextField = New-Object System.Windows.Forms.TextBox
    $SharePathSingleTextField.Text = "$RepoShareCommon"
    $SharePathSingleTextField.Multiline = $false 
    $SharePathSingleTextField.location = New-Object System.Drawing.Point(($LeftOffset+80),($TopOffset+15)) # (from left, from top)
    $SharePathSingleTextField.Size = New-Object System.Drawing.Size(380,$FieldHeight)             # (width, height)
    $SharePathSingleTextField.ReadOnly = $true
    $SharePathSingleTextField.Name = "Single_Share_Path"
    #$SharePathSingleTextField.BorderStyle = 'None'                                           # 'none', 'FixedSingle', 'Fixed3D (default)'
    
    $ShareButton = New-Object System.Windows.Forms.RadioButton
    $ShareButton.Location = '10,14'
    $CommonButton = New-Object System.Windows.Forms.RadioButton
    $CommonButton.Location = '10,34'

    $lBackgroundColor = 'LightSteelBlue'

    # populate CM Udate checkbox from .INI variable setting
    $find = "^[\$]CommonRepo"
    (Get-Content $IniFIleFullPath) | Foreach-Object { 
            if ($_ -match $find) { 
                if ( $_ -match '\$true' ) { 
                    $CommonButton.Checked = $true                          # set the visual default from the INI setting
                    $SharePathSingleTextField.BackColor = $lBackgroundColor
                } else { 
                    $ShareButton.Checked = $true 
                    $SharePathTextField.BackColor = $lBackgroundColor
                }
            } # if ($_ -match $find)
        } # Foreach-Object

    # update .INI variable setting - $CommonRepo 
    $CommonRepoRadio_Click = {
        $find = "^[\$]CommonRepo"

        if ( $CommonButton.checked ) {
            $Script:CommonRepo = $true
            $SharePathSingleTextField.BackColor = $lBackgroundColor
            $SharePathTextField.BackColor = ""
        } else {
            $Script:CommonRepo = $false
            $SharePathSingleTextField.BackColor = ""
            $SharePathTextField.BackColor = $lBackgroundColor        
        }
        $replace = "`$CommonRepo = `$$Script:CommonRepo"                   # set up the replacing string to either $false or $true from ini file

        Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"

        get_filters $dataGridView $false                                   # re-populate categories for available systems
        
    } # $CommonRepoRadio_Click

    $PathsGroupBox = New-Object System.Windows.Forms.GroupBox
    $PathsGroupBox.location = New-Object System.Drawing.Point(($LeftOffset+355),($TopOffset-10)) # (from left, from top)
    $PathsGroupBox.Size = New-Object System.Drawing.Size(($FormWidth-415),60)                    # (width, height)
    $PathsGroupBox.text = "Repository Paths - from $($IniFile):"

    $ShareButton.Add_Click( $CommonRepoRadio_Click )
    $CommonButton.Add_Click( $CommonRepoRadio_Click )                

    $PathsGroupBox.Controls.AddRange(@($SharePathLabel, $SharePathTextField, $SharePathLabelSingle, $SharePathSingleTextField, $ShareButton, $CommonButton, $HPIAPathLabel, $HPIAPathField))

    $CM_form.Controls.AddRange(@($PathsGroupBox))

    #----------------------------------------------------------------------------------
    # Create CM checkboxes GroupBoxes
    #----------------------------------------------------------------------------------

    $CMGroupBox = New-Object System.Windows.Forms.GroupBox
    $CMGroupBox.location = New-Object System.Drawing.Point($LeftOffset,($TopOffset+50))         # (from left, from top)
    $CMGroupBox.Size = New-Object System.Drawing.Size(($FormWidth-650),40)                       # (width, height)
    #$CMGroupBox.FlatStyle = 'System'                                                              # Flat, Popup, Standard, System
    $CMGroupBox.Text = 'SCCM - Disconnected'

    $CMGroupBox2 = New-Object System.Windows.Forms.GroupBox
    $CMGroupBox2.location = New-Object System.Drawing.Point(($LeftOffset+260),($TopOffset+50))         # (from left, from top)
    $CMGroupBox2.Size = New-Object System.Drawing.Size(($FormWidth-320),40)                       # (width, height)


    #----------------------------------------------------------------------------------
    # Create CM Repository Packages Update button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Repo Checkbox' }
    $updateCMCheckbox = New-Object System.Windows.Forms.CheckBox
    $updateCMCheckbox.Text = 'Update Packages /'
    $updateCMCheckbox.Autosize = $true
    $updateCMCheckbox.Location = New-Object System.Drawing.Point(($LeftOffset),($TopOffset-5))   # (from left, from top)

    # populate CM Udate checkbox from .INI variable setting - $UpdateCMPackages
    $find = "^[\$]UpdateCMPackages"
    (Get-Content $IniFIleFullPath) | Foreach-Object { 
        if ($_ -match $find) { 
            if ( $_ -match '\$true' ) { 
                $updateCMCheckbox.Checked = $true 
            } else { 
                $updateCMCheckbox.Checked = $false 
            }              } 
        } # Foreach-Object

    $Script:UpdateCMPackages = $updateCMCheckbox.Checked

    $updateCMCheckbox_Click = {
        $find = "^[\$]UpdateCMPackages"
        if ( $updateCMCheckbox.checked ) {
            $Script:UpdateCMPackages = $true
        } else {
            $Script:UpdateCMPackages = $false
            $Script:DistributeCMPackages = $false
            $CMDistributeheckbox.Checked = $false
        }
        $replace = "`$UpdateCMPackages = `$$Script:UpdateCMPackages"                   # set up the replacing string to either $false or $true from ini file
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

    # populate CM Udate checkbox from .INI variable setting - $DistributeCMPackages
    $find = "^[\$]DistributeCMPackages"
    (Get-Content $IniFIleFullPath) | Foreach-Object { 
        if ($_ -match $find) { 
            if ( $_ -match '\$true' ) { 
                $CMDistributeheckbox.Checked = $true 
            } else { 
                $CMDistributeheckbox.Checked = $false 
            }
        } # if ($_ -match $find)
        } # Foreach-Object

    $Script:DistributeCMPackages = $DistributeCMPackages.Checked

    $CMDistributeheckbox_Click = {
        $find = "^[\$]DistributeCMPackages"
        if ( $CMDistributeheckbox.checked ) {
            $Script:DistributeCMPackages = $true
        } else {
            $Script:DistributeCMPackages = $false
        }
        $replace = "`$DistributeCMPackages = `$$Script:DistributeCMPackages"                   # set up the replacing string to either $false or $true from ini file
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
            $find = "^[\$]HPIAVersion"
            $replace = "`$HPIAVersion = '$($Script:HPIAVersion)'"       
            Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"
            $find = "^[\$]HPIAPath"
            $replace = "`$HPIAPath = '$($Script:HPIAPath)'"       
            Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"
            $HPIAPathField.Text = "$HPIACMPackage - $HPIAPath"
            CM_HPIAPackage $HPIACMPackage $HPIAPath $HPIAVersion
        }
    } # $HPIAPathButton_Click = 

    $HPIAPathButton.add_Click($HPIAPathButton_Click)

    $HPIAPathField = New-Object System.Windows.Forms.TextBox
    $HPIAPathField.Text = "$HPIACMPackage - $HPIAPath"
    $HPIAPathField.Multiline = $false 
    $HPIAPathField.location = New-Object System.Drawing.Point(($LeftOffset+100),($TopOffset-5)) # (from left, from top)
    $HPIAPathField.Size = New-Object System.Drawing.Size(320,$FieldHeight)                      # (width, height)
    $HPIAPathField.ReadOnly = $true
    $HPIAPathField.Name = "HPIAPath"
    $HPIAPathField.BorderStyle = 'none'                                                  # 'none', 'FixedSingle', 'Fixed3D (default)'

    #----------------------------------------------------------------------------------
    # Create HPIA Browse button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating HPIA Browse button' }
    $buttonBrowse = New-Object System.Windows.Forms.Button
    $buttonBrowse.Width = 60
    $buttonBrowse.Text = 'Browse'
    $buttonBrowse.Location = New-Object System.Drawing.Point(($LeftOffset+480),($TopOffset-10))
    $buttonBrowse_Click={
        $FileBrowser.InitialDirectory = $HPIAPath
        $FileBrowser.Title = "Browse folder for HPImageAssistant.exe"
        $FileBrowser.Filter = "exe file (*Assistant.exe) | *assistant.exe" 
        $lBrowsePath = $FileBrowser.ShowDialog()    # returns 'OK' or Cancel'
        if ( $lBrowsePath -eq 'OK' ) {
            $lLeafHPIAExeName = Split-Path $FileBrowser.FileName -leaf          
            if ( $lLeafHPIAExeName -match 'hpimageassistant.exe' ) {
                $Script:HPIAPath = Split-Path $FileBrowser.FileName
                $Script:HPIAVersion = (Get-Item $FileBrowser.FileName).versioninfo.fileversion
                $find = "^[\$]HPIAVersion"
                $replace = "`$HPIAVersion = '$($Script:HPIAVersion)'"       
                Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"
                $find = "^[\$]HPIAPath"
                $replace = "`$HPIAPath = '$($Script:HPIAPath)'"       
                Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"
                $HPIAPathField.Text = "$HPIACMPackage - $HPIAPath"
                CMTraceLog -Message "... HPIA Path now: ''$HPIAPath'' (ver.$HPIAVersion) - May want to update SCCM package"
            } else {
                CMTraceLog -Message "... HPIA Path [$HPIAPath] does not contain HPIA executable"
            }
           # write-host $lLeafHPIAName $lNewHPIAPath 
        }
    } # $buttonBrowse_Click={
    $buttonBrowse.add_Click($buttonBrowse_Click)

    $CMGroupBox2.Controls.AddRange(@($HPIAPathButton, $HPIAPathField, $buttonBrowse))
    
    $CM_form.Controls.AddRange(@($CMGroupBox, $CMGroupBox2))

    #----------------------------------------------------------------------------------
    # Create Models list Checked Grid box - add 1st checkbox column
    # The ListView control allows columns to be used as fields in a row
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating DataGridView' }
    $ListViewWidth = ($FormWidth-80)
    $ListViewHeight = 250
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.location = New-Object System.Drawing.Point(($LeftOffset-10),($TopOffset))
    $dataGridView.height = $ListViewHeight
    $dataGridView.width = $ListViewWidth
    $dataGridView.ColumnHeadersVisible = $true                   # the column names becomes row 0 in the datagrid view
    $dataGridView.ColumnHeadersHeightSizeMode = 'AutoSize'       # AutoSize, DisableResizing, EnableResizing
    $dataGridView.RowHeadersVisible = $false
    $dataGridView.SelectionMode = 'CellSelect'
    $dataGridView.AllowUserToAddRows = $False                    # Prevents the display of empty last row
    if ( $DebugMode ) {  Write-Host 'creating col 0 checkbox' }
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
    $CheckAll.AutoSize=$true
    $CheckAll.Left=9
    $CheckAll.Top=10
    $CheckAll.Checked = $false

    $CheckAll_Click={

        $state = $CheckAll.Checked
        if ( $state ) {
            for ($i = 0; $i -lt $dataGridView.RowCount; $i++) {
                $dataGridView[0,$i].Value = $state  
                $dataGridView.Rows[$i].Cells['Driver'].Value = $state
                $dataGridView.Rows[$i].Cells['INI Software'].Value = $state
            }
        } else {
            clear_datagrid $dataGridView
        }

    } # $CheckAll_Click={

    $CheckAll.add_Click($CheckAll_Click)
    
    $dataGridView.Controls.Add($CheckAll)
    
    #----------------------------------------------------------------------------------
    # add columns 1, 2 (0 is 1st column)
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'adding SysId, Model columns' }
    $dataGridView.ColumnCount = 3                                # 1st column (0) is checkbox column

    $dataGridView.Columns[1].Name = 'SysId'
    $dataGridView.Columns[1].Width = 40
    $dataGridView.Columns[1].DefaultCellStyle.Alignment = "MiddleCenter"

    $dataGridView.Columns[2].Name = 'Model'
    $dataGridView.Columns[2].Width = 220

    #################################################################

    #----------------------------------------------------------------------------------
    # Add checkbox columns for every category filter
    # from column 4 on (0 is 1st column)
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating category columns' }
    foreach ( $cat in $FilterCategories ) {
        $temp = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
        $temp.name = $cat
        $temp.width = 50
        [void]$DataGridView.Columns.Add($temp) 
    }

    # populate with all the HP Models listed in the ini file
    $HPModelsTable | ForEach-Object {
                        # populate 1st 3 columns: checkmark, ProdId, Model Name
                        $row = @( $true, $_.ProdCode, $_.Model)         
                        [void]$dataGridView.Rows.Add($row)
                } # ForEach-Object
    
    #----------------------------------------------------------------------------------
    # add an All INI Software column
    # column 7 (0 is 1st column)
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating INI Software column' }
    $CheckBoxINISoftware = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $CheckBoxINISoftware.Name = 'INI Software' #'INI S/W'
    $CheckBoxINISoftware.width = 50
   
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
    $dataGridView.CurrentCell = $dataGridView[1,1]
    $dataGridView.ClearSelection()

    #----------------------------------------------------------------------------------
    # handle ALL checkbox selections here
    # columns= 0(row selected), 1(ProdCode), 2(Prod Name), 3-6(categories, 7(INI Software), 8(repository path)
    #----------------------------------------------------------------------------------
    $CellOnClick = {

        $row = $this.currentRow.index 
        $column = $this.currentCell.ColumnIndex 

        # checkmark cell returns a 'bool' $true or $false, so review those clicks
        $CurrCellValue = $dataGridView.rows[$row].Cells[$column].value
        
        # next, see if the cell is a checkmark (type Boolean) or a text cell (which would NOT have a value of $true or $false
        if ( $dataGridView.rows[$row].Cells[$column].Value.GetType() -eq [Boolean] ) {

            $CellprevState = $dataGridView.rows[$row].Cells[$column].EditedFormattedValue # value 'BEFORE' click action $true/$false
            $CellNewState = !$CellprevState

            # here we know we are dealing with one of the checkmark/selection cells

            switch ( $column ) {

                0 {                                                               # need to reset all categories
                    if ( $CellNewState ) {
                        # default to check; Driver and ini.ps1 Software list for this model
                        $dataGridView.Rows[$row].Cells['Driver'].Value = $true
                        $dataGridView.Rows[$row].Cells['INI Software'].Value = $true
                    } else {                               
                        $datagridview.Rows[$row].Cells['INI Software'].Value = $false      # reset the 'All' column
                        foreach ( $cat in $FilterCategories ) {                   # ... all categories          
                            $datagridview.Rows[$row].Cells[$cat].Value = $false
                        } # forech ( $cat in $FilterCategories )
                        $datagridview.Rows[$row].Cells[$datagridview.columnCount-1].Value = '' # ... and reset the repo path field
                    } # if ( $CellnewState )
                } # 0

                Default {                                                         
                    # here to deal with clicking on a category cell or 'INI Software'
                    if ( -not $CellNewState ) {
                        foreach ( $cat in $FilterCategories ) {
                            $CatColumn = $datagridview.Rows[$row].Cells[$cat].ColumnIndex
                            if ( $CatColumn -eq $column ) {
                                continue                                              
                            } else {
                                # see if anouther category column is checked
                                if ( $datagridview.Rows[$row].Cells[$cat].Value ) {
                                    $CellNewState = $true
                                }
                            } # else if ( $colClicked -eq $currColumn )
                        } # foreach ( $cat in $FilterCategories )

                        # confirm if INI Software is selected/checked 
                        if ( ($column -ne 7) -and ($datagridview.Rows[$row].Cells['INI Software'].Value) ) {
                            $CellNewState = $true
                        }
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
    $CMModelsGroupBox.location = New-Object System.Drawing.Point($LeftOffset,($TopOffset+90))     # (from left, from top)
    $CMModelsGroupBox.Size = New-Object System.Drawing.Size(($ListViewWidth+20),($ListViewHeight+30))       # (width, height)
    $CMModelsGroupBox.text = "HP Models / Repository Category Filters"

    $CMModelsGroupBox.Controls.AddRange(@($dataGridView))

    $CM_form.Controls.AddRange(@($CMModelsGroupBox))
    
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
    $TextBox.location = New-Object System.Drawing.Point($LeftOffset,($FormHeight-380))            # (from left, from top)
    $TextBox.Size = New-Object System.Drawing.Size(($FormWidth-60),($FormHeight/2-80))             # (width, height)

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
    # Add a Refresh Grid button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating a Refresh Grid button' }
    $RefreshGridButton = New-Object System.Windows.Forms.Button
    #$RefreshGridButton.Location = New-Object System.Drawing.Point(($LeftOffset+100),($FormHeight-47))    # (from left, from top)
    $RefreshGridButton.Location = New-Object System.Drawing.Point(($LeftOffset),($FormHeight-410))    # (from left, from top)
    $RefreshGridButton.Text = 'Refresh Grid'
    $RefreshGridButton.AutoSize=$true

    $RefreshGridButton_Click={
        Get_Filters  $dataGridView $false
    } # $RefreshGridButton_Click={

    $RefreshGridButton.add_Click($RefreshGridButton_Click)

    $CM_form.Controls.Add($RefreshGridButton)
  
    #----------------------------------------------------------------------------------
    # Add a list filters button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating a list filters button' }
    $ListFiltersdButton = New-Object System.Windows.Forms.Button
    $ListFiltersdButton.Location = New-Object System.Drawing.Point(($LeftOffset+100),($FormHeight-410))    # (from left, from top)
    $ListFiltersdButton.Text = 'List Filters'
    $ListFiltersdButton.AutoSize=$true

    $ListFiltersdButton_Click={
        Get_Filters  $dataGridView $true
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
    $NewLogCheckbox.Location = New-Object System.Drawing.Point(($LeftOffset+360),($FormHeight-403))

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
    $LogPathLabel.Location = New-Object System.Drawing.Point(($LeftOffset+410),($FormHeight-402))    # (from left, from top)
    $LogPathLabel.AutoSize=$true

    $LogPathField = New-Object System.Windows.Forms.TextBox
    $LogPathField.Text = "$LogFile"
    $LogPathField.Multiline = $false 
    $LogPathField.location = New-Object System.Drawing.Point(($LeftOffset+440),($FormHeight-405)) # (from left, from top)
    $LogPathField.Size = New-Object System.Drawing.Size(400,$FieldHeight)                      # (width, height)
    $LogPathField.ReadOnly = $true
    $LogPathField.Name = "LogPath"
    #$LogPathField.BorderStyle = 'none'                                                  # 'none', 'FixedSingle', 'Fixed3D (default)'
    # next, move cursor to end of text in field, to see the log file name
    $LogPathField.Select($LogPathField.Text.Length,0)
    $LogPathField.ScrollToCaret()

    $CM_form.Controls.AddRange(@($LogPathLabel, $LogPathField ))
  
    #----------------------------------------------------------------------------------
    # Create 'Debug Mode' - checkmark
    #----------------------------------------------------------------------------------
    $DebugCheckBox = New-Object System.Windows.Forms.CheckBox
    $DebugCheckBox.Text = 'Debug Mode'
    $DebugCheckBox.UseVisualStyleBackColor = $True
    $DebugCheckBox.location = New-Object System.Drawing.Point(($LeftOffset+100),($FormHeight-45))   # (from left, from top)
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
    # Add a TextBox larger and smaller Font buttons
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating a TextBox smaller Font button' }
    $TextBoxFontdButtonDec = New-Object System.Windows.Forms.Button
    $TextBoxFontdButtonDec.Location = New-Object System.Drawing.Point(($LeftOffset+200),($FormHeight-45))    # (from left, from top)
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
    $TextBoxFontdButtonInc.Location = New-Object System.Drawing.Point(($LeftOffset+280),($FormHeight-45))    # (from left, from top)
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
    if ( $DebugMode ) { Write-Host 'creating Done Button' }
    $buttonDone = New-Object System.Windows.Forms.Button
    $buttonDone.Text = 'Done'
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

    #----------------------------------------------------------------------------------
    # Finally, show the dialog on screen
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'calling get_filters' }
    get_filters $dataGridView $false

    if ( $DebugMode ) { Write-Host 'calling ShowDialog' }
    $CM_form.ShowDialog() | Out-Null

} # Function CreateForm 

# --------------------------------------------------------------------------
# Start of Script
# --------------------------------------------------------------------------

# at this point, we are past the -h-help runstring option... if any left, then we need to run without the UI

# in case we need to browse for a file, create the object now
	
Add-Type -AssemblyName System.Windows.Forms
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
            $CommonRepo = $true } else { $CommonRepo = $false }  } 
    if ( $PSBoundParameters.Keys.Contains('Products') ) { "-Products: $($Products)" }
    if ( $PSBoundParameters.Keys.Contains('ListFilters') ) { list_filters }
    if ( $PSBoundParameters.Keys.Contains('NoIniSw') ) { '-NoIniSw' }
    if ( $PSBoundParameters.Keys.Contains('showActivityLog') ) { $showActivityLog = $true } 
    if ( $PSBoundParameters.Keys.Contains('Sync') ) { sync_repos }

} # if ( $MyInvocation.BoundParameters.count -gt 0)

