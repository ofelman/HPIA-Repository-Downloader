<#
    HP Image Assistant Softpaq Repository Downloader
    by Dan Felman/HP Technical Consultant
        
        Loop Code: The 'HPModelsTable' loop code, is in a separate INI.ps1 file 
            ... created by Gary Blok's (@gwblok) post on garytown.com.
            ... https://garytown.com/create-hp-bios-repository-using-powershell
 
        Logging: The Log function based on code by Ryan Ephgrave (@ephingposh)
            ... https://www.ephingadmin.com/powershell-cmtrace-log-function/
        Version informtion in Release_Notes.txt
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

$ScriptVersion = "2.00.00 (April 7, 2022)"

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

# with v2.00, we introduce $v_OPSYS variable to handle win10/11 OS versions
if ( Get-Variable -Name v_OPSYS -ErrorAction SilentlyContinue ) {
    "Using updated INI file to support Win11, (v_OPSYS=$Script:v_OPSYS)"
} else {
    'Need updated INI file (required to support Win11)' ; exit
}
#--------------------------------------------------------------------------------------
# add required .NET framwork items to support GUI
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

#--------------------------------------------------------------------------------------
# Script Environment Vars

$CMConnected = $false                                # is a connection to SCCM established?
$SiteCode = $null

$s_ModelAdded = $False                               # set $True when user adds models to list in the UI
$s_AddSoftware = '.ADDSOFTWARE'                      # sub-folders where named downloaded Softpaqs will reside
$s_HPIAActivityLog = 'activity.log'                    # name of HPIA activity log file
#$s_HPServerIPFile = "$($ScriptPath)\15"                # used to temporarily hold IP connections, while job is running

$MyPublicIP = (Invoke-WebRequest ifconfig.me/ip).Content.Trim()
'My Public IP: '+$MyPublicIP | Out-Host

#--------------------------------------------------------------------------------------
# error codes for color coding, etc.
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

    if ( ($Type -ne $TypeDebug) -or ( ($Type -eq $TypeDebug) -and $v_DebugMode) ) {
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

    if ( $v_DebugMode ) { CMTraceLog -Message "> Load_HPCMSLModule" -Type $TypeNorm }
    $m = 'HPCMSL'

    CMTraceLog -Message "Checking for required HP CMSL modules... " -Type $TypeNoNewline

    # If module is imported say that and do nothing
    if (Get-Module | Where-Object {$_.Name -eq $m}) {
        if ( $v_DebugMode ) { write-host "Module $m is already imported." }
        CMTraceLog -Message "Module already imported." -Type $TypSuccess
    } else {

        # If module is not imported, but available on disk then import
        if (Get-Module -ListAvailable | Where-Object {$_.Name -eq $m}) {
            if ( $v_DebugMode ) { write-host "Importing Module $m." }
            CMTraceLog -Message "Importing Module $m." -Type $TypeNoNewline
            Import-Module $m -Verbose
            CMTraceLog -Message "Done" -Type $TypSuccess
        } else {

            # If module is not imported, not available on disk, but is in online gallery then install and import
            if (Find-Module -Name $m | Where-Object {$_.Name -eq $m}) {
                if ( $v_DebugMode ) { write-host "Upgrading NuGet and updating PowerShellGet first." }
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
                if ( $v_DebugMode ) { write-host "Installing and Importing Module $m." }
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

    if ( $v_DebugMode ) { CMTraceLog -Message "< Load_HPCMSLModule" -Type $TypeNorm }

} # function Load_HPCMSLModule

#=====================================================================================
<#
    Function Test_CMConnection
        The function will test the CM server connection
        and that the Task Sequences required for use of the Script are available in CM
        - will also test that both download and share paths exist
#>
Function Test_CMConnection {

    if ( $v_DebugMode ) { CMTraceLog -Message "> Test_CMConnection" -Type $TypeNorm }

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
    if ( $v_DebugMode ) { CMTraceLog -Message "< Test_CMConnection" -Type $TypeNorm }

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

        if ( $v_DebugMode ) { CMTraceLog -Message "... setting location to: $($SiteCode)" -Type $TypeDebug }
        Set-Location -Path "$($SiteCode):"

        #--------------------------------------------------------------------------------
        # now, see if HPIA package exists
        #--------------------------------------------------------------------------------

        $lCMHPIAPackage = Get-CMPackage -Name $pHPIAPkgName -Fast

        if ( $null -eq $lCMHPIAPackage ) {
            if ( $v_DebugMode ) { CMTraceLog -Message "... HPIA Package missing... Creating New" -Type $TypeNorm }
            $lCMHPIAPackage = New-CMPackage -Name $pHPIAPkgName -Manufacturer "HP"
            CMTraceLog -ErrorMessage "... HPIA package created - PackageID $lCMHPIAPackage.PackageId"
        } else {
            CMTraceLog -Message "... HPIA package found - PackageID $($lCMHPIAPackage.PackageId)"
        }
        #if ( $v_DebugMode ) { CMTraceLog -Message "... setting HPIA Package to Version: $($Script:v_OSVER), path: $($pRepoPath)" -Type $TypeDebug }

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

    if ( $v_DebugMode ) {  CMTraceLog -Message "> CM_RepoUpdate" -Type $TypeNorm }

    # develop the Package name
    $lPkgName = 'HP-'+$pModelProdId+'-'+$pModelName
    CMTraceLog -Message "... updating repository for SCCM package: $($lPkgName)" -Type $TypeNorm

    if ( $v_DebugMode ) { CMTraceLog -Message "... setting location to: $($SiteCode)" -Type $TypeDebug }
    Set-Location -Path "$($SiteCode):"

    if ( $v_DebugMode ) { CMTraceLog -Message "... getting CM package: $($lPkgName)" -Type $TypeDebug }
    $lCMRepoPackage = Get-CMPackage -Name $lPkgName -Fast

    if ( $null -eq $lCMRepoPackage ) {
        CMTraceLog -Message "... Package missing... Creating New" -Type $TypeNorm
        $lCMRepoPackage = New-CMPackage -Name $lPkgName -Manufacturer "HP"
    }
    #--------------------------------------------------------------------------------
    # update package with info from share folder
    #--------------------------------------------------------------------------------
    if ( $v_DebugMode ) { CMTraceLog -Message "... setting CM Package to Version: $($Script:v_OSVER), path: $($pRepoPath)" -Type $TypeDebug }

	Set-CMPackage -Name $lPkgName -Version "$($Script:v_OSVER)"
	Set-CMPackage -Name $lPkgName -Path $pRepoPath

    if ( $Script:v_DistributeCMPackages  ) {
        CMTraceLog -Message "... updating CM Distribution Points"
        update-CMDistributionPoint -PackageId $lCMRepoPackage.PackageID
    }
    $lCMRepoPackage = Get-CMPackage -Name $lPkgName -Fast                               # make sure we are woring with updated/distributed package

    #--------------------------------------------------------------------------------

    Set-Location -Path $pCurrentLoc

    if ( $v_DebugMode ) { CMTraceLog -Message "< CM_RepoUpdate" -Type $TypeNorm }

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
        $lRepoLogFile = "$($ldotRepository)\$($s_HPIAActivityLog)"

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
        ... and initialize it for HPIA, if arg $pInitialize = $True
        Args:
            $pRepoFOlder: folder to validate, or create
            $pInitRepository: $true, initialize repository, 
                              $false: do not initialize (used as root of individual repository folders)

#>
Function init_repository {
    [CmdletBinding()]
	param( $pRepoFolder,
            $pInitialize )

    if ( $v_DebugMode ) { CMTraceLog -Message "> init_repository" -Type $TypeNorm }

    $lCurrentLoc = Get-Location

    $retRepoCreatedFlag = $false

    if ( (Test-Path $pRepoFolder) -and ($pInitialize -eq $false) ) {
        $retRepoCreatedFlag = $true
    } else {
        $lRepoParentPath = Split-Path -Path $pRepoFolder -Parent        # Get the Parent path to the Main Repo folder

        #--------------------------------------------------------------------------------
        # see if we need to create the path to the folder

        if ( !(Test-Path $lRepoParentPath) ) {
            Try {
                # create the Path to the Repo folder
                $lret = New-Item -Path $lRepoParentPath -ItemType directory
            } Catch {
                CMTraceLog -Message "[init_repository] Done" -Type $TypeNorm
                return $retRepoCreatedFlag
            } # Catch
        } # if ( !(Test-Path $lRepoPathSplit) ) 

        #--------------------------------------------------------------------------------
        # now add the Repo folder if it doesn't exist
        if ( !(Test-Path $pRepoFolder) ) {
            CMTraceLog -Message "... creating Repository Folder $pRepoFolder" -Type $TypeNorm
            $lret = New-Item -Path $pRepoFolder -ItemType directory
        } # if ( !(Test-Path $pRepoFolder) )

        $retRepoCreatedFlag = $true

        #--------------------------------------------------------------------------------
        # if needed, check on repository to initialize (CMSL repositories have a .Repository folder)

        if ( $pInitialize -and !(test-path "$pRepoFolder\.Repository")) {
            Set-Location $pRepoFolder
            $initOut = (Initialize-Repository) 6>&1
            CMTraceLog -Message  "... Repository Initialization done $($Initout)"  -Type $TypeNorm 

            CMTraceLog -Message  "... Configuring $($pRepoFolder) for HP Image Assistant" -Type $TypeNorm
            Set-RepositoryConfiguration -setting OfflineCacheMode -cachevalue Enable 6>&1   # configuring the repo for HP IA's use
            # configuring to create 'Contents.CSV' after every Sync
            Set-RepositoryConfiguration -setting RepositoryReport -Format csv 6>&1   
        } # if ( $pInitialize -and !(test-path "$pRepoFOlder\.Repository"))

    } # else if ( (Test-Path $pRepoFolder) -and ($pInitialize -eq $false) )

    #--------------------------------------------------------------------------------
    # intialize .ADDSOFTWARE folder for holding named softpaqs
    # ... this folder will hold softpaqs added outside CMSL' sync/cleanup root folder
    #--------------------------------------------------------------------------------
    $lAddSoftpaqsFolder = "$($pRepoFolder)\$($s_AddSoftware)"

    if ( !(Test-Path $lAddSoftpaqsFolder) -and $pInitialize ) {
        CMTraceLog -Message "... creating Add-on Softpaq Folder $lAddSoftpaqsFolder" -Type $TypeNorm
        $lret = New-Item -Path $lAddSoftpaqsFolder -ItemType directory
        if ( $v_DebugMode ) { CMTraceLog -Message "NO $lAddSoftpaqsFolder" -Type $TypeWarn }
    } # if ( !(Test-Path $lAddSoftpaqsFolder) )

    if ( $v_DebugMode ) { CMTraceLog -Message "< init_repository" -Type $TypeNorm }

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
        Requires: pGrid - entries from the UI grid devices
                  pFolder - repository folder to check out
                  pCommonRepoFlag - check for using Common repo for all devices
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

    switch ( $pCommonRepoFlag ) {
        #--------------------------------------------------------------------------------
        # Download AddOns for every device with flag files content in the repo
        #--------------------------------------------------------------------------------
        $true { 
            #--------------------------------------------------------------------------------
            # There are likely multiple models being held in this repository
            # ... so find all Platform AddOns flag files for content defined by user
            #--------------------------------------------------------------------------------
            for ( $iRow = 0; $iRow -lt $pGrid.RowCount; $iRow++ ) {
        
                $lProdCode = $pGrid[1,$iRow].Value                      # column 1 = Model/Prod ID
                $lModelName = $pGrid[2,$iRow].Value                     # column 2 = Model name
                $lAddOnsFlag = $pGrid.Rows[$iRow].Cells['AddOns'].Value # column 7 = 'AddOns' checkmark
                $lProdIDFlagFile = $lAddSoftpaqsFolder+'\'+$lProdCode

                if ( $lAddOnsFlag -and (Test-Path $lProdIDFlagFile) ) {

                    if ( $v_DebugMode ) { CMTraceLog -Message 'Flag file found:'+$lAddOnsFlagFile -Type $TypeWarn }
                    [array]$lAddOnsList = Get-Content $lProdIDFlagFile

                    if ( -not ([String]::IsNullOrWhiteSpace($lAddOnsList)) ) {
                        CMTraceLog -Message "... platform: $($lProdCode): checking AddOns flag file"
                        if ( $v_DebugMode ) { CMTraceLog -Message 'calling Get-SoftpaqList():'+$lProdCode -Type $TypeNorm }
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
                                        CMTraceLog -Message "... done" -Type $TypeNorm
                                    } # else if ( Test-Path $lSoftpaqExe ) {
                                    CMTraceLog -Message "`tdownloading $($iSoftpaq.id) CVA file" -Type $TypeNoNewline
                                    $ret = (Get-SoftpaqMetadataFile $iSoftpaq.Id) 6>&1   # ALWAYS update the CVA file, even if previously downloaded
                                    CMTraceLog -Message "... done" -Type $TypeNorm
                                } # if ( [string]$iSoftpaq.Name -match $iEntry )
                            } # ForEach ( $iSoftpaq in $lSoftpaqList )
                            if ( -not $lEntryFound ) {
                                CMTraceLog -Message  "... '$($iEntry)': Softpaq not found for this platform and OS version"  -Type $TypeWarn
                            } # if ( -not $lEntryFound )
                        } # ForEach ( $lEntry in $lAddOnsList )

                    } else {
                        CMTraceLog -Message $lProdCode': Flag file found but empty: '
                    } # else if ( -not ([String]::IsNullOrWhiteSpace($lAddOns)) )

                } else {
                    CMTraceLog -Message '... '$lProdCode': AddOns cell not checked, will not attempt to download'
                } # else if (Test-Path $lAddOnsFlagFile)

            } # for ( $i = 0; $i -lt $pGrid.RowCount; $i++ )
        } # true

        #--------------------------------------------------------------------------------
        # Download AddOns for an individual model ( $pCommonRepoFlag = $false )
        #--------------------------------------------------------------------------------
        $false {
            #--------------------------------------------------------------------------------
            # Search grid for one model's product ID, so we can find the AddOns flag file
            #--------------------------------------------------------------------------------
            for ( $i = 0; $i -lt $pGrid.RowCount; $i++ ) {
                $lSystemID = $pGrid[1,$i].Value                          # column 1 has the Model/Prod ID
                $lSystemName = $pGrid[2,$i].Value                         # column 2 has the Model name
                $lFoldername = split-path $pFolder -Leaf
                if ( $lFoldername -match $lSystemName ) {                 # we found the mode, so get the Prod ID
                    $lAddOnsValue = $pGrid.Rows[$i].Cells['AddOns'].Value # column 7 has the AddOns checkmark
                    $lProdIDFlagFile = $lAddSoftpaqsFolder+'\'+$lSystemID

                    if ( $lAddOnsValue -and (Test-Path $lProdIDFlagFile) ) {

                        if ( $v_DebugMode ) { CMTraceLog -Message 'Flag file found:'+$lProdIDFlagFile -Type $TypeWarn }
                        [array]$lAddOnsList = Get-Content $lProdIDFlagFile
                        if ( -not ([String]::IsNullOrWhiteSpace($lAddOnsList)) ) {
                            CMTraceLog -Message "... found AddOns in flag file for platform: $($lSystemID.id)"
                            $lSoftpaqList = (Get-SoftpaqList -platform $lSystemID -os $Script:OS -osver $Script:v_OSVER -characteristic SSM) 6>&1  

                            ForEach ( $iEntry in $lAddOnsList ) {
                                ForEach ( $iSoftpaq in $lSoftpaqList ) {
                                    if ( [string]$iSoftpaq.Name -match $iEntry ) {
                                        
                                        $lSoftpaqExe = $lAddSoftpaqsFolder+'\'+$iSoftpaq.id+'.exe'
                                        if ( Test-Path $lSoftpaqExe ) {
                                            CMTraceLog -Message "`t$($iSoftpaq.id) already downloaded - $($iSoftpaq.Name)"
                                        } else {
                                            CMTraceLog -Message "`tfound $($iSoftpaq.id) $($iSoftpaq.Name)" -Type $NoNewLine
                                            $ret = (Get-Softpaq $iSoftpaq.Id) 6>&1
                                            CMTraceLog -Message  "- Downloaded $($ret)"  -Type $TypeWarn
                                        } # else if ( Test-Path $lSoftpaqExe ) {
                                        # ALWAYS update the CVA file, even if previously downloaded
                                        $ret = (Get-SoftpaqMetadataFile $iSoftpaq.Id) 6>&1
                                        if ( $v_DebugMode ) { CMTraceLog -Message  "`tGet-SoftpaqMetadataFile done: $($ret)"  -Type $TypeWarn }
                                    } # if ( [string]$iSoftpaq.Name -match $iEntry )
                                } # ForEach ( $iSoftpaq in $lSoftpaqList )
                            } # ForEach ( $lEntry in $lAddOnsList )

                        } else {
                            CMTraceLog -Message $lSystemID'Flag file found but empty'
                        } # else if ( -not ([String]::IsNullOrWhiteSpace($lAddOns)) )
                    } # if ( $lAddOnsValue -and (Test-Path $lProdIDFlagFile) )
                    
                } # if ( $lFoldername -match $lSystemName )
            } # for ( $i = 0; $i -lt $pGrid.RowCount; $i++ )
        } # false

    } #  switch ( $pCommonRepoFlag )
    
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
            CMTraceLog -Message "... Restoring named softpaqs/cva files to Repository"
            $lsource = "$($pFolder)\$($s_AddSoftware)\*.*"
            Copy-Item -Path $lsource -Filter *.CVA -Destination $pFolder
            Copy-Item -Path $lsource -Filter *.EXE -Destination $pFolder
            CMTraceLog -Message "... Restoring named softpaqs/cva files completed"
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

        #--------------------------------------------------------------------------------
        CMTraceLog -Message  '... calling invoke-repositorysync()' -Type $TypeNorm
        invoke-repositorysync 6>&1
        #--------------------------------------------------------------------------------

        # find what sync'd from the CMSL log file for this run
        try {
            $lRepoFilters = (Get-RepositoryInfo).Filters  # see if any filters are used to continue
            Get_ActivityLogEntries $pFolder  # get sync'd info from HPIA activity log file
            $lContentsHASH = (Get-FileHash -Path "$($pFolder)\.repository\contents.csv" -Algorithm MD5).Hash
            if ( $v_DebugMode ) { CMTraceLog -Message "... MD5 Hash of 'contents.csv': $($lContentsHASH)" -Type $TypeWarn }
            #--------------------------------------------------------------------------------
            CMTraceLog -Message  '... calling invoke-RepositoryCleanup()' -Type $TypeNorm
            invoke-RepositoryCleanup 6>&1

            #-----------------------------------------------------------------------------------
            # Now, we manage the 'AddOns' Softpaqs maintained in .ADDSOFTWARE
            # ... each platform should have a flag file w/ID name 
            #-----------------------------------------------------------------------------------
            CMTraceLog -Message  '... calling Download_AddOn_Softpaqs()' -Type $TypeNorm
            Download_AddOn_Softpaqs $pGrid $pFolder $pCommonFlag

            CMTraceLog -Message  '... calling Restore_Softpaqs()' -Type $TypeNorm
            # next, copy all softpaqs in $s_AddSoftware subfolder to the repository 
            # ... (since it got cleared up by CMSL's "Invoke-RepositoryCleanup")
            Restore_AddOn_Softpaqs $pFolder (-not $script:noIniSw)
        } catch {
            $lContentsHASH = $null
        }
        #--------------------------------------------------------------------------------
        # see if Cleanup modified the contents.csv file 
        # - seems like (up to at least 1.6.3) RepositoryCleanup does not Modify 'contents.csv'
        #$lContentsHASH = Get_SyncdContents $pFolder       # Sync command creates 'contents.csv'
        #CMTraceLog -Message "... MD5 Hash of 'contents.csv': $($lContentsHASH)" -Type $TypeWarn
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
    if ( $v_DebugMode ) { CMTraceLog -Message "... adding category filters" -Type $TypeDebug }

    foreach ( $cat in $Script:v_FilterCategories ) {
        if ( $pGrid.Rows[$pRow].Cells[$cat].Value ) {
            CMTraceLog -Message  "... adding filter: -Platform $pModelID -os $OPSYS -osver $Script:v_OSVER -category $($cat) -characteristic ssm" -Type $TypeNoNewline
            $lRes = (Add-RepositoryFilter -platform $pModelID -os $Script:v_OS  -osver $Script:v_OSVER -category $cat -characteristic ssm 6>&1)
            CMTraceLog -Message $lRes -Type $TypeWarn 
        }
    } # foreach ( $cat in $Script:v_FilterCategories )
    #--------------------------------------------------------------------------------
    # update repository path field for this model in the grid (path is the last col)
    #--------------------------------------------------------------------------------
    $pGrid[($pGrid.ColumnCount-1),$pRow].Value = $pModelRepository

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
            $pNewModels )                                 # $True = models added to list)                                      # array of rows selected
    
    CMTraceLog -Message "> sync_individual_repositories - START" -Type $TypeNorm
    if ( $v_DebugMode ) { CMTraceLog -Message "... for OS version: $($Script:v_OSVER)" -Type $TypeDebug }

    #--------------------------------------------------------------------------------
    # decide if we work on single repo or individual repos
    #--------------------------------------------------------------------------------     
    init_repository $pRepoHeadFolder $Null #$false         # make sure Main repo folder exists, do NOT make it a repository

    #--------------------------------------------------------------------------------
    if ( $v_DebugMode ) { CMTraceLog -Message "... stepping through all selected models" -Type $TypeDebug }

    # go through every selected HP Model in the list
    for ( $i = 0; $i -lt $pGrid.RowCount; $i++ ) {
        
        $lModelId = $pGrid[1,$i].Value                    # column 1 has the Model/Prod ID
        $lModelName = $pGrid[2,$i].Value                  # column 2 has the Model name

        # if model entry is checked, we need to create a repository, but ONLY if $Script:v_CommonRepo = $false
        if ( $i -in $pCheckedItemsList ) {

            CMTraceLog -Message "`n--- Updating model: $lModelName"
            $lTempRepoFolder = "$($pRepoHeadFolder)\$($lModelId)_$($lModelName)"     # this is the repo folder for this model

            if ( -not (Test-Path -Path $lTempRepoFolder) ) {
                $lTempRepoFolder = "$($pRepoHeadFolder)\$($lModelName)"           # this is the repo folder for this model, without SysID in name
            }

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

    if  ( ($pCheckedItemsList).count -eq 0 ) { return }

    $pCurrentLoc = Get-Location

    CMTraceLog -Message "> Sync_Common_Repository - START" -Type $TypeNorm
    if ( $v_DebugMode ) { CMTraceLog -Message "... for OS version: $($Script:v_OSVER)" -Type $TypeDebug }

    init_repository $pRepoFolder $true        # make sure Main repo folder exists, or create it, init
    CMTraceLog -Message  "... Common repository selected: $($pRepoFolder)" -Type $TypeNorm

    if ( $Script:v_KeepFilters ) {
        CMTraceLog -Message  "... Keeping existing filters in: $($pRepoFolder)" -Type $TypeNorm
    } else {
        CMTraceLog -Message  "... Removing existing filters in: $($pRepoFolder)" -Type $TypeNorm
    } # else if ( $Script:v_KeepFilters )

    if ( $v_DebugMode ) { CMTraceLog -Message "... stepping through selected models" -Type $TypeDebug }
    
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
            if ( ((Get-RepositoryInfo).Filters).count -gt 0 ) {
                $lres = (Remove-RepositoryFilter -platform $lModelID -yes 6>&1)
                if ( $v_DebugMode ) { CMTraceLog -Message "... removed filters for: $($lModelID)" -Type $TypeWarn }
            }
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

    if ( $v_DebugMode ) { CMTraceLog -Message '> clear_grid' -Type $TypeNorm }

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
    
    if ( $v_DebugMode ) { CMTraceLog -Message '< clear_grid' -Type $TypeNorm }

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
        $lproducts = $lProdFilters.platform | Get-Unique

        # develop the filter list by product
        foreach ( $iProduct in $lproducts ) {
            $lcharacteristics = $lProdFilters.characteristic | Get-Unique 
            $lOS = $lProdFilters.operatingSystem | Get-Unique
            $lMsg = "   -Platform $($iProduct) -OS $($lOS) -Category"
            foreach ( $iFilter in $lProdFilters ) {
                if ( $iFilter.Platform -eq $iProduct ) { $lMsg += ' '+$iFilter.category }
            } # foreach ( $iFilter in $lProdFilters )
            $lMsg += " -Characteristic $($lcharacteristics)"
            CMTraceLog -Message $lMsg -Type $TypeNorm
        }
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

        foreach ( $iprodName in $lProdFolders ) {
            set-location "$($Script:v_Root_IndividualRepoFolder)\$($iprodName.name)"

            $lProdFilters = (get-repositoryinfo).Filters
            $lprods = $lProdFilters.platform | Get-Unique 
            $lcharacteristics = $lProdFilters.characteristic | Get-Unique            
            $lOS = $lProdFilters.operatingSystem | Get-Unique 

            $lMsg = "   -Platform $($lprods) -OS $($lOS) -Category"
            foreach ( $icat in $lProdFilters.category ) { $lMsg += ' '+$icat }
            $lMsg += " -Characteristic $($lcharacteristics)"
            CMTraceLog -Message $lMsg -Type $TypeNorm
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

    if ( $v_DebugMode ) { CMTraceLog -Message '> Get_CommonRepofilters' -Type $TypeNorm }

    # see if the repository was configured for HPIA
    if ( !(Test-Path "$($pCommonFolder)\.Repository") ) {
        CMTraceLog -Message "... Repository Folder not initialized" -Type $TypeWarn
        if ( $v_DebugMode ) { CMTraceLog -Message '< Get_CommonRepofilters - Done' -Type $TypeWarn }
        return
    } 

    if ( $pRefreshGrid ) {
        CMTraceLog -Message "... Refreshing Grid from Common Repository: ''$pCommonFolder'' (calling clear_grid())" -Type $TypeNorm
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
                            CMTraceLog -Message "... $lPlatform - Additional Softpaqs: $lAddOns" -Type $TypeWarn  
                        } #                    
                    } # if ( Test-Path $lAddOnRepoFile )
                } # if ( $lPlatform -eq $pGrid[1,$i].value )
            } else {
                CMTraceLog -Message "... listing filters (calling List_Filters())" -Type $TypeWarn   
                List_Filters
            } # else if ( $pRefreshGrid )
        } # for ( $i = 0; $i -lt $pGrid.RowCount; $i++ )
    } # foreach ( $filter in $lProdFilters )

    [void]$pGrid.Refresh
    CMTraceLog -Message "... Refreshing Grid ...DONE" -Type $TypeSuccess

    if ( $v_DebugMode ) { CMTraceLog -Message '< Get_CommonRepofilters()' -Type $TypeNorm }
    
} # Function Get_CommonRepofilters

#=====================================================================================
<#
    Function Get_IndividualRepofilters
        Retrieves category filters from the repository for each selected model
        ... and populates the Grid appropriately
        Parameters:
            $pGrid                          The models grid in the GUI
            $pRepoLocation                      Where to start looking for repositories
#>
Function Get_IndividualRepofilters {
    [CmdletBinding()]
	param( $pGrid,                                  # array of row lines that are checked
        $pRepoRoot )

    if ( $v_DebugMode ) { CMTraceLog -Message '> Get_IndividualRepofilters' -Type $TypeNorm }

    set-location $pRepoRoot

    CMTraceLog -Message '... Refreshing Grid from Individual Repositories ...' -Type $TypeNorm

    #--------------------------------------------------------------------------------
    # now check for each product's repository folder
    # if the repo is created, then check the category filters
    #--------------------------------------------------------------------------------
    for ( $iRow = 0; $iRow -lt $pGrid.RowCount; $iRow++ ) {

        $lModelId = $pGrid[1,$iRow].Value                               # column 1 has the Model/Prod ID
        $lModelName = $pGrid[2,$iRow].Value                             # column 2 has the Model name 
        $lTempRepoFolder = "$($pRepoRoot)\$($lModelId)_$($lModelName)"  # this is the repo folder for this model
        if ( -not (Test-Path -Path $lTempRepoFolder) ) {
            $lTempRepoFolder = "$($pRepoRoot)\$($lModelName)"           # this is the repo folder for this model, without SysID in name
        }

        if ( Test-Path $lTempRepoFolder ) {
            set-location $lTempRepoFolder                               # move to location of Repository
                                    
            $pGrid[($dataGridView.ColumnCount-1),$iRow].Value = $lTempRepoFolder
            
            # if we are maintaining AddOns Softpaqs, show it in the checkbox
            $lAddOnRepoFile = $lTempRepoFolder+'\'+$s_AddSoftware+'\'+$lModelId   
            if ( Test-Path $lAddOnRepoFile ) {      
                if ( [String]::IsNullOrWhiteSpace((Get-content $lAddOnRepoFile)) ) {
                            $pGrid.Rows[$iRow].Cells['AddOns'].Value = $False
                } else {
                    $pGrid.Rows[$iRow].Cells['AddOns'].Value = $True
                    [array]$lAddOns = Get-Content $lAddOnRepoFile
                    $lMsg = "... Additional Softpaqs Enabled for Platform '$lModelId' {$lAddOns}"
                    CMTraceLog -Message $lMsg -Type $TypeWarn 
                }
            } # if ( Test-Path $lAddOnRepoFile )

            $lProdFilters = (get-repositoryinfo).Filters
            
            # platform        : 8715
            # operatingSystem : win10:21h1
            # category        : BIOS
            # releaseType     : *
            # characteristic  : ssm

            $lplatformList = @()
            foreach ( $iEntry in $lProdFilters ) {
                if ( $v_DebugMode ) { CMTraceLog -Message "... Platform $($lProdFilters.platform) ... $($lProdFilters.operatingSystem) $($lProdFilters.category) $($lProdFilters.characteristic) - @$($lTempRepoFolder)" -Type $TypeWarn }
                if ( -not ($iEntry.platform -in $lplatformList) ) {
                    $lplatformList += $iEntry.platform
                    if ( $v_DebugMode ) { CMTraceLog -Message "... populating filter categories ''$($lModelName)''" -Type $TypeNorm }
                    foreach ( $cat in  ($lProdFilters.category.split(' ')) ) {
                        $pGrid.Rows[$iRow].Cells[$cat].Value = $true
                    }
                } # else 
            } # foreach ( $iEntry in $lProdFilters )
        } # if ( Test-Path $lTempRepoFolder )
    } # for ( $iRow = 0; $iRow -lt $pModelsList.RowCount; $iRow++ ) 

    if ( $v_DebugMode ) { CMTraceLog -Message '< Get_IndividualRepofilters]' -Type $TypeNorm }
    
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
        Browse to find an existing, or create new, repository for individual platforms
        Paramaters: $pGrid = pointer to model grid in UI
                    $pCurrentRepository (the path listed in UI) - start as the default path
#>
Function Browse_IndividualRepos {
    [CmdletBinding()]
	param( $pGrid, $pCurrentRepository )  

    CMTraceLog -Message ">> Browse Individual Repositories" -Type $TypeSuccess

    # let's find the repository to import by asking the user

    $lbrowse = New-Object System.Windows.Forms.FolderBrowserDialog
    $lbrowse.SelectedPath = $pCurrentRepository         
    $lbrowse.Description = "Select a Root Folder for individual repositories"
    $lbrowse.ShowNewFolderButton = $true
                                  
    if ( $lbrowse.ShowDialog() -eq "OK" ) {
        $lRepository = $lbrowse.SelectedPath
        CMTraceLog -Message  "... clearing list (calling Empty_Grid()" -Type $TypeNorm
        Empty_Grid $pGrid
        Set-Location $lRepository
        CMTraceLog -Message  "... checking for existing repositories at: $($lRepository)" -Type $TypeNorm
        $lDirectories = (Get-ChildItem -Directory)     
        # Let's search for repositories (subfolders) at this location
        # add a row in the DataGrid for every repository (e.g. model) we find
        $lReposFound = $False     # assume no repositories exist here
        foreach ( $lFolder in $lDirectories ) {
            $lRepoFolder = $lRepository+'\'+$lFolder
            Set-Location $lRepoFolder
            Try {
                $lRepoFilters = (Get-RepositoryInfo).Filters
                CMTraceLog -Message  "... [Repository Found ($($lRepoFolder)) -- adding to grid]" -Type $TypeNorm
                # obtain platform SysID from filters in the repository
                [array]$lRepoPlatform = $lRepoFilters.platform
                [void]$pGrid.Rows.Add(@( $true, $lRepoPlatform, $lFolder ))
                $lReposFound = $True
            } Catch {
                CMTraceLog -Message  "... $($lRepoFolder) is NOT a Repository" -Type $TypeNorm
            } # Catch
        } # foreach ( $lFolder in $lDirectories )
        
        if ( $lReposFound ) {
            CMTraceLog -Message  "... Retrieving Filters from repositories (calling Get_IndividualRepofilters())"
            Get_IndividualRepofilters $pGrid $lRepository
            if ( $v_DebugMode ) { CMTraceLog -Message  "... calling Update_INIModelsList()" }
            Update_INIModelsList $pGrid $lRepository $False  # $False = treat as head of individual repositories
        } else {
            CMTraceLog -Message  "... no repositories found"
            $lResponse = [System.Windows.MessageBox]::Show('Do you want to clear the HPModel list?','INI File update','YesNo')
            if ( $lResponse -eq 'Yes' ) {
                CMTraceLog -Message  "... Clearing HPModels lists calling Update_UIandINI()"
                Update_UIandINI $lRepository $False              # $False = using individual repositories
            } else {
                CMTraceLog -Message  "... HPModels list not cleared"
            }
        } # else if ( $lReposFound )
    } else {
        $lRepository = $pCurrentRepository
    } # else if ( $lbrowse.ShowDialog() -eq "OK" )
    
    $lbrowse.Dispose()

    CMTraceLog -Message "<< Browse Individual Repositories Done ($($lRepository))" -Type $TypeSuccess

    Return $lRepository

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

    CMTraceLog -Message "> Import_IndividualRepos at ($($pCurrentRepository))" -Type $TypeNorm

    if ( -not (Test-Path $pCurrentRepository) ) {
        CMTraceLog "Individual Root Path from INI file not found: $($pCurrentRepository) - will create"  -Type $TypeWarn
        Init_Repository $pCurrentRepository $False            # False = create root folder only
    }
    CMTraceLog -Message  "... clearing Grid (calling Empty_Grid())" -Type $TypeNorm
    Empty_Grid $pGrid
    Set-Location $pCurrentRepository

    $lDirectories = (Get-ChildItem -Directory)                    # each subfolder is an Individual model repository

    if ( $lDirectories.count -eq 0 ) {
        #--------------------------------------------------------------------------------
        # if the repository has no subfolders, populate models from the INI's @HPModels list
        #--------------------------------------------------------------------------------        
        CMTraceLog -Message  "... (calling Populate_Grid_from_INI())" -Type $TypeNorm
        Populate_Grid_from_INI $pGrid $pINIModelsList
    } else {
        #--------------------------------------------------------------------------------
        # else populate from each valid individual repository subfolder
        #--------------------------------------------------------------------------------
        foreach ( $ientry in $lDirectories ) {
            $lRepoFolderFullPath = $pCurrentRepository+'\'+$ientry
            $lRepoAddSoftwareFolderFullPath = $lRepoFolderFullPath+"\.ADDSOFTWARE"
            $lRepoAddSoftwareFlagFile = Get-ChildItem $lRepoAddSoftwareFolderFullPath | Where-Object {$_.BaseName.Length -eq 4} # flag file name has 4 characters
            #Write-Host ($lRepoAddSoftwareFlagFile).count $ientry

            # handle repos created prior to using SysID in folder name
            If ( [string]$ientry.name -match "^[0-9].*" ) {
                $lProdName = $ientry.name.substring(5) # if folder name starts with an Integer, it has the SysID in the name, so pick model name after SysID
                $lentrySysID = $ientry.name.substring(0,4)  # platform may have >1 motherboard IDs
            } else {
                $lProdName = $ientry.name
                if ( ($lRepoAddSoftwareFlagFile).count -eq 0 ) {
                    $lentrySysID = (Get-HPDeviceDetails -Name $lProdName).SystemID  # platform may have >1 motherboard IDs
                } else {
                    $lentrySysID = $lRepoAddSoftwareFlagFile
                }
            } # else If ([string]$ientry.name -match "^[0-9].*")           

            Set-Location $lRepoFolderFullPath

            Try {
                CMTraceLog -Message  "... Checking repository: $($lRepoFolderFullPath) -- adding to grid" -Type $TypeNorm
                $lRepoFilters = (Get-RepositoryInfo).Filters | Get-Unique
                if ( $lRepoFilters.count -eq 0 ) {
                    # get the platform ID from the AddOns flag file... this folder should have a single flag file
                    $lPlatformID = Get-ChildItem $lRepoAddSoftwareFolderFullPath
                    $lRow = $pGrid.Rows.Add(@( $False, $lPlatformID.Name.Substring(0,4), $lProdName ))
                } else {
                    
                    $lRow = $pGrid.Rows.Add(@( $True, $lentrySysID, $lProdName ))
                } # else if ( $lRepoFilters.count -eq 0 )
                # --------------------------------------------------------------
                # now check for the AddOns flag file, and if it has content
                # --------------------------------------------------------------
                $lAddOnsFlagFile = $lRepoAddSoftwareFolderFullPath+'\'+$lentrySysID
                if ( Test-Path -Path $lAddOnsFlagFile -PathType leaf ) {
                    $lisFlagFileEmpty = [String]::IsNullOrWhiteSpace((Get-content $lAddOnsFlagFile))
                    if ( $lisFlagFileEmpty ) {
                        CMTraceLog -Message  "... [flag file empty (no addons):$lAddOnsFlagFile]" -Type $TypeNorm
                        $pGrid.rows[$lRow].Cells['AddOns'].Value = $False
                    } else {
                        $pGrid.rows[$lRow].Cells['AddOns'].Value = $True
                    }
                } else {
                    CMTraceLog -Message  "... [flag file missing:$lAddOnsFlagFile]" -Type $TypeWarn
                } # else if ( Test-Path -Path $lAddOnsFlagFile )
            } Catch {
                CMTraceLog -Message  "... NOT a Repository: ($($lRepoFolderFullPath))" -Type $TypeNorm
            } # Catch
        } # foreach ( $ientry in $lDirectories )

        CMTraceLog -Message  "... Getting Filters from individual repositories (Get_IndividualRepofilters())"
        Get_IndividualRepofilters $pGrid $pCurrentRepository
        CMTraceLog -Message  "... Updating UI and INI Path (calling Update_UIandINI())"
        Update_UIandINI $pCurrentRepository $False    # $False = individual repos
        CMTraceLog -Message  "... Updating UI HPModels list (calling Update_INIModelsList())"
        Update_INIModelsList $pGrid $pCurrentRepository $False  # $False = this is for Individual repositories
    } # else if ( $lDirectories.count -eq 0 )

    CMTraceLog -Message '< Import_IndividualRepos done' -Type $TypeNorm

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
                $row = @( $false, $lRepoPlatforms[$i], $lProdName[0].Name )  
                [void]$pGrid.Rows.Add($row)
            } # for ( $i = 0 ; $i -lt $lRepoPlatforms.count ; $i++)
            CMTraceLog -Message "... Finding filters from $($lRepository) (calling Get_CommonRepofilters())" -Type $TypeNorm
            Get_CommonRepofilters $pGrid $lRepository $True
            CMTraceLog -Message "... found HPIA repository $($lRepository)" -Type $TypeNorm
        } Catch {
            # lets initialize repository but maintain existing models in the grid
            # ... in case we want to use some of them in the new repo
            init_repository $lRepository $true
            clear_grid $pGrid
            CMTraceLog -Message "... $($lRepository) initialized " -Type $TypeNorm
        } # catch
        Update_INIModelsList $pGrid $lRepository $True  # $True = this is a common repository
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
        CMTraceLog -Message  "... looking for filters in repository" -Type $TypeNorm

        if ( $lProdFilters.count -eq 0 ) {
            CMTraceLog -Message  "... no filters found, so populate from INI file (calling Populate_Grid_from_INI())" -Type $TypeNorm
            Populate_Grid_from_INI $pGrid $pModelsTable
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
<#            
            # find flag files in the repo without filters applied
            # ... meaning, model was added, flag file created, but not filters are selected
            $lAddSoftwarePath = $pCurrentRepository+'\.ADDSOFTWARE'
            $lFlagFiles = Get-ChildItem $lAddSoftwarePath -Recurse -File | Where-Object {$_.BaseName.Length -eq 4} #| Measure-Object | %{$_.Count}# 
            foreach ( $iPlatform in $lFlagFiles ) {
                write-host $iPlatform
                if ( $iPlatform -notin $lRepoPlatforms ) {
                    write-host $iPlatform' has no filters'
                    [array]$lProdName = Get-HPDeviceDetails -Platform $iPlatform
                    [void]$pGrid.Rows.Add(@( $false, $iPlatform, $lProdName[0].Name ))
                }
            } # foreach ( $iPlatform in $lFlagFiles )
#>
            CMTraceLog -Message "... Finding filters from $($pCurrentRepository) (calling Get_CommonRepofilters())" -Type $TypeNorm
            Get_CommonRepofilters $pGrid $pCurrentRepository $True

            CMTraceLog -Message  "... Updating UI and INI file from filters (calling Update_UIandINI())" -Type $TypeNorm
            Update_UIandINI $pCurrentRepository $True     # $True = common repo

            CMTraceLog -Message  "... Updating models in INI file (calling Update_INIModelsList())" -Type $TypeNorm
            Update_INIModelsList $pGrid $pCurrentRepository $True  # $False means treat as head of individual repositories
        } # if ( $lProdFilters.count -gt 0 )
        CMTraceLog -Message "< Import Repository() Done" -Type $TypeSuccess
    } Catch {
        CMTraceLog -Message "< Repository Folder ($($pCurrentRepository)) not initialized for HPIA" -Type $TypeWarn
    } # Catch

} # Function Import_CommonRepo

#=====================================================================================
<#
    Function Check_PlatformsOSVersion
    here we check the OS version, if supported by any platform selected in list
#>
Function Check_PlatformsOSVersion  {
    [CmdletBinding()]
	param( $pGrid,
            $pOS,
            $pOSVersion )

    if ( $v_DebugMode ) { CMTraceLog -Message '> Check_PlatformsOSVersion]' -Type $TypeNorm }
    
    CMTraceLog -Message "Checking support for $($pOS)/$($pOSVersion) for selected platforms" -Type $TypeNorm

    # search thru the table entries for checked items, and see if each product
    # has support for the selected OS version

    for ( $i = 0; $i -lt $pGrid.RowCount; $i++ ) {

        if ( $pGrid[0,$i].Value ) {
            $lPlatform = $pGrid[1,$i].Value
            $lPlatformName = $pGrid[2,$i].Value

            # get list of OS versions supported by this platform and OS version
            $l_OSList = get-hpdevicedetails -platform $lPlatform -OSList -ErrorAction Continue

            $l_OSIsSupported = $false
            foreach ( $entry in $l_OSList ) {
                if ( ($pOSVersion -eq $entry.OperatingSystemRelease) -and $entry.OperatingSystem -match $pOS.substring(3) ) {
                    CMTraceLog -Message  "... $pOS/$pOSVersion OS is supported for $($lPlatform)/$($lPlatformName)" -Type $TypeNorm
                    $l_OSIsSupported = $true
                }
            } # foreach ( $entry in $lOSList )
            if ( -not $l_OSIsSupported ) {
                CMTraceLog -Message  "... $pOS/$pOSVersion OS is NOT supported for $($lPlatform)/$($lPlatformName)" -Type $TypeWarn
            }
        } # if ( $dataGridView[0,$i].Value )  

    } # for ( $i = 0; $i -lt $pGrid.RowCount; $i++ )
    
    if ( $v_DebugMode ) { CMTraceLog -Message '< Check_PlatformsOSVersion]' -Type $TypeNorm }

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

    if ( $v_DebugMode ) { CMTraceLog -Message '> Update_INIModelsList' -Type $TypeNorm }

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
        CMTraceLog -Message "... Updating Models list (INI File $lModelsListINIFile)" -Type $TypeNorm
        (get-content $lModelsListINIFile) -replace $lListHeader, "$&$($lModelsList)" | 
            Set-Content $lModelsListINIFile
    } else {
        CMTraceLog -Message " ... INI file not updated - didn't find" -Type $TypeWarn
    } # else if ( Test-Path $lRepoLogFile )

    if ( $v_DebugMode ) { CMTraceLog -Message '< Update_INIModelsList' -Type $TypeNorm }
    
} # Function Update_INIModelsList 

#=====================================================================================
<#
    Function Populate_Grid_from_INI
    This is the MAIN function with a Gui that sets things up for the user
#>
Function Populate_Grid_from_INI {
    [CmdletBinding()]
	param( $pGrid,
            $pModelsTable )
<#
    populate with all the HP Models listed in the ini file
    excample line: 
    @{ ProdCode = '8723'; Model = 'HP ZBook Firefly 14 G7 Mobile Workstation' }
#>
    CMTraceLog -Message "Populating Grid from INI file's `$HPModels list" -Type $TypeNorm
    $pModelsTable | 
        ForEach-Object {
            [void]$pGrid.Rows.Add( @( $False, $_.ProdCode, $_.Model) )   # populate checkmark, ProdId, Model Name
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
        $IndividualPathTextField.BackColor = ""
        $Script:v_Root_CommonRepoFolder = $pNewPath
        $find = "^[\$]v_Root_CommonRepoFolder"
        $replace = "`$v_Root_CommonRepoFolder = ""$pNewPath"""
    } else {
        $IndividualRadioButton.Checked = $True
        $IndividualPathTextField.Text = $pNewPath
        $IndividualPathTextField.BackColor = $BackgroundColor
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
    Function Manage_AddOnFlag_File
        Creates and updates the Pltaform ID flag file in the repository's .ADDSOFTWARE folder
        If Flag file exists and user 'unselects' AddOns, we move the file to a backup file 'ABCD_bk'
    Parameters                 
        $pPath
        $pSysID
        $pModelName
        [array]$pAddOns
        $pCreateFile -- flag                # $True to create the flag file, $False to remove (renamed)
#>
Function Manage_AddOnFlag_File {
    [CmdletBinding()]
	param( $pPath,
            $pSysID,                        # 4-digit hex-code platform/motherboard ID
            $pModelName,
            [array]$pAddOns,                # $Null means empty flag file (nothing to add)
            $pCreateFile,                   # $True = create file, $False = rename file
            $pCommonFlag) 
                              
    if ( $pCommonFlag ) {
        $lFlagFilePath = $pPath+'\'+$s_AddSoftware+'\'+$pSysID
    } else {
        $lRpoPath = $pPath+'\'+$pModelName
        if ( -not (Test-Path $lRpoPath) ) {
            $lRpoPath = $pPath+'\'+$pSysID+'_'+$pModelName
        }
        $lFlagFilePath = $lRpoPath+'\'+$s_AddSoftware+'\'+$pSysID       
    } # else if ( $pCommonFlag )

    $lFlagFilePathBK = $lFlagFilePath+'_BK'

    if ( $pCreateFile ) {
        $lMsg = "... $pSysID - Enabling download of AddOns Softpaqs "
        if ( -not (Test-Path $lFlagFilePath ) ) { 
            if ( Test-Path $lFlagFilePathBK ) { 
                Move-Item $lFlagFilePathBK $lFlagFilePath -Force 
                [array]$lAddOnsList = Get-Content $lFlagFilePath
                if ( -not ([String]::IsNullOrWhiteSpace($lAddOnsList)) ) { $lMsg += ": ... $lAddOnsList" }
            } else {
                New-Item $lFlagFilePath 
                if ( $pAddOns.count -gt 0 ) { 
                    $pAddOns[0] | Out-File -FilePath $lFlagFilePath 
                    For ($i=1; $i -le $pAddOns.count; $i++) { $pAddOns[$i] | Out-File -FilePath $lFlagFilePath -Append }
                } # if ( $pAddOns.count -gt 0 )
            } # else if ( Test-Path $lFlagFilePath'_BK' )
        } # if ( -not (Test-Path $lFlagFilePath ) )
        CMTraceLog -Message "... Selecting additional Softpaqs for model $pModelName (calling EditAddonFlag())" -Type $TypeNorm
        EditAddonFlag $lFlagFilePath
    } else {
        $lMsg = "... $pSysID - Disabling download of AddOns Softpaqs"

        if ( Test-Path $lFlagFilePath ) { Move-Item $lFlagFilePath $lFlagFilePathBK -Force }
    } # else if ( $pCreateFile )

    # find out how many entries in the flag file
    if ( Test-Path $lFlagFilePath ) {
        [int]$lLines = (Get-Content $lFlagFilePath | Measure-Object).Count
    } else {
        [int]$lLines = (Get-Content $lFlagFilePathBK | Measure-Object).Count
    }

    CMTraceLog -Message $lMsg -Type $TypeNorm
    return $lLines

} # Manage_AddOnFlag_File

#=====================================================================================
<#
    Function EditAddonFlag

    manage the Softpaq Addons in the flag file
    called to confirm what entries to be maintained in the ID flag file
#>
Function EditAddonFlag { 
    [CmdletBinding()]
	param( $pFlagFilePath )

    $pSysID = Split-Path $pFlagFilePath -leaf
    $pFileContents = Get-Content $pFlagFilePath

    $fEntryFormWidth = 400
    $fEntryFormHeigth = 400
    $fOffset = 20
    $FieldHeight = 20
    $fPathFieldLength = 200

    if ( $v_DebugMode ) { Write-Host 'Manage AddOn Entries Form' }
    $SoftpaqsForm = New-Object System.Windows.Forms.Form
    $SoftpaqsForm.MaximizeBox = $False ; $SoftpaqsForm.MinimizeBox = $False #; $EntryForm.ControlBox = $False
    $SoftpaqsForm.Text = "Additional Softpaqs"
    $SoftpaqsForm.Width = $fEntryFormWidth ; $SoftpaqsForm.height = 400 ; $SoftpaqsForm.Autosize = $true
    $SoftpaqsForm.StartPosition = 'CenterScreen'
    $SoftpaqsForm.Topmost = $true

    # -----------------------------------------------------------
    $EntryId = New-Object System.Windows.Forms.Label
    $EntryId.Text = "Name"
    $EntryId.location = New-Object System.Drawing.Point($fOffset,$fOffset) # (from left, from top)
    $EntryId.Size = New-Object System.Drawing.Size(60,20)                   # (width, height)

    $SoftpaqEntry = New-Object System.Windows.Forms.TextBox
    $SoftpaqEntry.Text = ""
    $SoftpaqEntry.Multiline = $false 
    $SoftpaqEntry.location = New-Object System.Drawing.Point(($fOffset+70),($fOffset-4)) # (from left, from top)
    $SoftpaqEntry.Size = New-Object System.Drawing.Size($fPathFieldLength,$FieldHeight)# (width, height)
    $SoftpaqEntry.ReadOnly = $False
    $SoftpaqEntry.Name = "Softpaq Name"


    $AddButton = New-Object System.Windows.Forms.Button
    $AddButton.Location = New-Object System.Drawing.Point(($fPathFieldLength+$fOffset+80),($fOffset-6))
    $AddButton.Size = New-Object System.Drawing.Size(75,23)
    $AddButton.Text = 'Add'

    $AddButton_AddClick = {
        if ( $SoftpaqEntry.Text ) {
            if ( $SoftpaqEntry.Text -in $pFileContents ) {
                CMTraceLog -Message '... item already included' -Type $TypeNorm
            } else {
                $EntryList.items.add($SoftpaqEntry.Text)
                $lentries = @()
                foreach ( $lEntry in $EntryList.items ) { $lentries += $lEntry }
                Set-Content $pFlagFilePath -Value $lentries      # reset the file with needed AddOns
            } # else if ( $SoftpaqEntry.Text -in $pFileContents )
        } else {
            CMTraceLog -Message '... nothing to add' -Type $TypeNorm 
        } # else if ( $EntryModel.Text )
    } # $SearchButton_AddClick =
    $AddButton.Add_Click( $AddButton_AddClick )

    # -----------------------------------------------------------

    $lListHeight = $fEntryFormHeigth/2-60
    $EntryList = New-Object System.Windows.Forms.ListBox
    $EntryList.Name = 'Entries'
    $EntryList.Autosize = $false
    $EntryList.location = New-Object System.Drawing.Point($fOffset,60)  # (from left, from top)
    $EntryList.Size = New-Object System.Drawing.Size(($fEntryFormWidth-60),$lListHeight) # (width, height)

    foreach ( $iSoftpaqName in $pFileContents ) { $EntryList.items.add($iSoftpaqName) }

    $EntryList.Add_Click({
        #$AddSoftpaqList.items.Clear()
        #foreach ( $iName in $pSoftpaqList ) { $AddSoftpaqList.items.add($iName) }
    })
    #$EntryList.add_DoubleClick({ return $EntryList.SelectedItem })
    
    # -----------------------------------------------------------
    $removeButton = New-Object System.Windows.Forms.Button
    $removeButton.Location = New-Object System.Drawing.Point(($fEntryFormWidth-120),($fEntryFormHeigth-200))
    $removeButton.Size = New-Object System.Drawing.Size(75,23)
    $removeButton.Text = 'Remove'
    $removeButton.add_Click({
        if ( $EntryList.SelectedItem ) {
            $EntryList.items.remove($EntryList.SelectedItem)
            $lentries = @()
            foreach ( $lEntry in $EntryList.items ) { $lentries += $lEntry }
            Set-Content $pFlagFilePath -Value $lentries      # reset the file with needed AddOns
        }
    }) # $removeButton.add_Click
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(($fEntryFormWidth-120),($fEntryFormHeigth-80))
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $SoftpaqsForm.AcceptButton = $okButton
    $SoftpaqsForm.Controls.AddRange(@($EntryId,$SoftpaqEntry,$AddButton,$EntryList,$SpqLabel,$AddSoftpaqList, $removeButton, $okButton))
    # -------------------------------------------------------------------
    # Ask the user what model to add to the list
    # -------------------------------------------------------------------
    $lResult = $SoftpaqsForm.ShowDialog()

} # Function EditAddonFlag

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
    $fFieldHeight = 20
    $fPathFieldLength = 200

    if ( $v_DebugMode ) { Write-Host 'Add Entry Form' }
    $EntryForm = New-Object System.Windows.Forms.Form
    $EntryForm.MaximizeBox = $False ; $EntryForm.MinimizeBox = $False #; $EntryForm.ControlBox = $False
    $EntryForm.Text = "Find a Device to Add"
    $EntryForm.Width = $fEntryFormWidth ; $EntryForm.height = 400 ; $EntryForm.Autosize = $true
    $EntryForm.StartPosition = 'CenterScreen'
    $EntryForm.Topmost = $true

    # ------------------------------------------------------------------------
    # find and add model entry
    # ------------------------------------------------------------------------
    $EntryId = New-Object System.Windows.Forms.Label
    $EntryId.Text = "Name"
    $EntryId.location = New-Object System.Drawing.Point($fOffset,$fOffset) # (from left, from top)
    $EntryId.Size = New-Object System.Drawing.Size(60,20)                   # (width, height)

    $EntryModel = New-Object System.Windows.Forms.TextBox
    $EntryModel.Text = ""       # start w/INI setting
    $EntryModel.Multiline = $false 
    $EntryModel.location = New-Object System.Drawing.Point(($fOffset+70),($fOffset-4)) # (from left, from top)
    $EntryModel.Size = New-Object System.Drawing.Size($fPathFieldLength,$fFieldHeight)# (width, height)
    $EntryModel.ReadOnly = $False
    $EntryModel.Name = "Model Name"
    $EntryModel.add_MouseHover($ShowHelp)
    $SearchButton = New-Object System.Windows.Forms.Button
    $SearchButton.Location = New-Object System.Drawing.Point(($fPathFieldLength+$fOffset+80),($fOffset-6))
    $SearchButton.Size = New-Object System.Drawing.Size(75,23)
    $SearchButton.Text = 'Search'
    $SearchButton_AddClick = {
        if ( $EntryModel.Text ) {
            $AddEntryList.Items.Clear()
            $lModels = Get-HPDeviceDetails -Like -Name $EntryModel.Text    # find all models matching entered text
            foreach ( $iModel in $lModels ) { 
                [void]$AddEntryList.Items.Add($iModel.SystemID+'_'+$iModel.Name) 
            }
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

    # ------------------------------------------------------------------------
    # find and add initial softpaqs to the selected model
    # ------------------------------------------------------------------------
    $SpqLabel = New-Object System.Windows.Forms.Label
    $SpqLabel.Text = "Select initial Addon Softpaqs" 
    $SpqLabel.location = New-Object System.Drawing.Point($fOffset,($lListHeight+60)) # (from left, from top)
    $SpqLabel.Size = New-Object System.Drawing.Size(70,60)                   # (width, height)

    $AddSoftpaqList = New-Object System.Windows.Forms.ListBox
    $AddSoftpaqList.Name = 'Softpaqs'
    $AddSoftpaqList.Autosize = $false
    $AddSoftpaqList.SelectionMode = 'MultiExtended'
    $AddSoftpaqList.location = New-Object System.Drawing.Point(($fOffset+70),($lListHeight+60))  # (from left, from top)
    $AddSoftpaqList.Size = New-Object System.Drawing.Size(($fEntryFormWidth-130),($lListHeight-40)) # (width, height)

    # ------------------------------------------------------------------------
    # show the dialog, and once user preses OK, add the model and create the flag file for addons
    # ------------------------------------------------------------------------
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(($fEntryFormWidth-120),($fEntryFormHeigth-80))
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(($fEntryFormWidth-200),($fEntryFormHeigth-80))
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::CANCEL

    $EntryForm.AcceptButton = $okButton
    $EntryForm.CancelButton = $cancelButton
    $EntryForm.Controls.AddRange(@($EntryId,$EntryModel,$SearchButton,$AddEntryList,$SpqLabel,$AddSoftpaqList, $cancelButton, $okButton))

    $lResult = $EntryForm.ShowDialog()

    if ($lResult -eq [System.Windows.Forms.DialogResult]::OK) {
        
        [array]$lSelectedEntry = Get-HPDeviceDetails -Like -Name $AddEntryList.SelectedItem

        $lSelectedModel = $AddEntryList.SelectedItem.substring(5)  # name is after 'SysID_'
        $lSelectedSysID = $AddEntryList.SelectedItem.substring(0,4)

        # see if model is already in the grid, and avoid it, otherwise, add it
        for ( $iRow = 0; $iRow -lt $pGrid.RowCount; $iRow++ ) {
            $lCurrEntrySysID = $pGrid.Rows[$iRow].Cells[1].value
            if ( $lCurrEntrySysID -like $lSelectedSysID ) {
                $lRes = [System.Windows.MessageBox]::Show("This model is already in the Grid","Add Model to Grid",0)    # 1 = "OKCancel" ; 4 = "YesNo"
                return $null
            }
        } # for ( $iRow = 0; $iRow -lt $pGrid.RowCount; $iRow++ )


        # add model to UI grid
        [void]$pGrid.Rows.Add( @( $False, $lSelectedSysID, $lSelectedModel) )

        # get a list of selected additional softpaqs to add
        $lSelectedSoftpaqs = $AddSoftpaqList.SelectedItems
        CMTraceLog -Message '... managing Flag file (calling Manage_AddOnFlag_File())' -Type $TypeNorm 
        $lnumEntries = Manage_AddOnFlag_File $pRepositoryFolder $lSelectedSysID $lSelectedModel $lSelectedSoftpaqs $True $pCommonFlag

        # 'Manage_AddOnFlag_File' returns an object, not an int... the 1st 2 entries are always in object, e.g., no addons
        if ( $lnumEntries.Length -gt 2 ) {
            $pGrid.Rows[($pGrid.RowCount-1)].Cells['AddOns'].Value = $True
        }
        CMTraceLog -Message '... Updating INI Models list (calling Update_INIModelsList())' -Type $TypeNorm 
        Update_INIModelsList $pGrid $pRepositoryFolder $pCommonFlag
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
            "OS_Version"     {$tip = "What Windows version to work with"}
            "Keep Filters"     {$tip = "Do NOT erase previous product selection filters"}
            "Continue on 404"  {$tip = "Continue Sync evern with Error 404, missing files"}
            "Individual Paths" {$tip = "Path to Head of Individual platform repositories"}
            "Common Path"      {$tip = "Path to Common/Shared platform repository"}
            "Models Table"     {$tip = "HP Models table to Sync repository(ies) to"}
            "Check All"        {$tip = "This check selects all Platforms and categories"}
            "Sync"             {$tip = "Syncronize repository for selected items from HP cloud"}
            'Use INI List'     {$tip = 'Reset the Grid from the INI file $HPModelsTable list'}
            'Filters'          {$tip = 'Show list of all current Repository filters'}
            'Add Model'        {$tip = 'Find and add a model to the current list in the Grid'}
        } # Switch ($this.name)
        $CMForm_tooltip.SetToolTip($this,$tip)
    } #end ShowHelp

    if ( $v_DebugMode ) { Write-Host 'creating Form' }
    $CM_form = New-Object System.Windows.Forms.Form
    $CM_form.Text = "HPIARepo_Downloader v$($ScriptVersion)"
    $CM_form.Width = $FormWidth
    $CM_form.height = $FormHeight
    $CM_form.Autosize = $true
    $CM_form.StartPosition = 'CenterScreen'

    #----------------------------------------------------------------------------------
    # Create OS and OS Version display fields - info from .ini file
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating OS and OSVer fields' }
    $OSList = New-Object System.Windows.Forms.ComboBox
    $OSList.Size = New-Object System.Drawing.Size(60,$FieldHeight)                  # (width, height)
    $OSList.Location  = New-Object System.Drawing.Point(($LeftOffset+10), ($TopOffset+4))
    $OSList.DropDownStyle = "DropDownList"
    $OSList.Name = "OS"
    # populate menu list from INI file's $v_OPSYS variable
    Foreach ($MenuItem in $Script:v_OPSYS) {
        [void]$OSList.Items.Add($MenuItem);
    } 
    $OSList.SelectedItem = $Script:v_OS
    $OSList.add_SelectedIndexChanged( {
            $Script:v_OS = $OSList.SelectedItem
            $OSVERComboBox.Items.Clear()
            switch ( $Script:v_OS ) {   # re-populate the OS Version list
                'Win10' { Foreach ($Version in $Script:v_OSVALID10) { [void]$OSVERComboBox.Items.Add($Version) } }
                'Win11' { Foreach ($Version in $Script:v_OSVALID11) { [void]$OSVERComboBox.Items.Add($Version) } }
            }
            $OSVERComboBox.SelectedItem = $Script:v_OSVER  # default to last OS Version   
            
            $find = "^[\$]v_OS\ "
            $replace = "`$v_OS = '$Script:v_OS'"                   # set up the replacing string to either $false or $true from ini file
            Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"

    } ) # $OSList.add_SelectedIndexChanged()

    $OSVERComboBox = New-Object System.Windows.Forms.ComboBox
    $OSVERComboBox.Size = New-Object System.Drawing.Size(60,$FieldHeight)                  # (width, height)
    $OSVERComboBox.Location  = New-Object System.Drawing.Point(($LeftOffset+80), ($TopOffset+4))
    $OSVERComboBox.DropDownStyle = "DropDownList"
    $OSVERComboBox.Name = "OS_Version"
    $OSVERComboBox.add_MouseHover($ShowHelp)
    # populate menu list from INI file
    switch ( $Script:v_OS ) {   # re-populate the OS Version list
        'Win10' { Foreach ($Version in $Script:v_OSVALID10) { [void]$OSVERComboBox.Items.Add($Version) } }
        'Win11' { Foreach ($Version in $Script:v_OSVALID11) { [void]$OSVERComboBox.Items.Add($Version) } }
    } 
    $OSVERComboBox.SelectedItem = $Script:v_OSVER 
    $OSVERComboBox.add_SelectedIndexChanged( {
        $Script:v_OSVER = $OSVERComboBox.SelectedItem
        Check_PlatformsOSVersion $dataGridView $Script:v_OS $Script:v_OSVER
        
        $find = "^[\$]v_OSVER"
        $replace = "`$v_OSVER = '$Script:v_OSVER'" 
        Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"

    } ) # $OSVERComboBox.add_SelectedIndexChanged()
    
    $CM_form.Controls.AddRange(@($OSList,$OSVERComboBox))

    #----------------------------------------------------------------------------------
    # Create Keep Filters checkbox
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating Keep Filters Checkbox' }
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
    if ( $v_DebugMode ) { Write-Host 'creating Continue on Error 404 Checkbox' }
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
    # messages used after switching repositories
    $OSMessage = 'Check and confirm the OS/OS Version selected for this Repository is correct'
    $OSHeader = 'Imported Repository'

    if ( $v_DebugMode ) { Write-Host 'creating Shared radio button' }
    $IndividualRadioButton = New-Object System.Windows.Forms.RadioButton
    $IndividualRadioButton.Location = '10,14'
    $IndividualRadioButton.Add_Click( {
            $Script:v_CommonRepo = $False
            $find = "^[\$]v_CommonRepo"
            $replace = "`$v_CommonRepo = `$$Script:v_CommonRepo" 
            Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"
            Import_IndividualRepos $dataGridView $IndividualPathTextField.Text $Script:HPModelsTable
            [System.Windows.MessageBox]::Show($OSMessage,$OSHeader,0)    # 0='OK', 1="OKCancel", 4="YesNo"
        }
    ) # $IndividualRadioButton.Add_Click()

    if ( $v_DebugMode ) { Write-Host 'creating Individual field label' }
    $SharePathLabel = New-Object System.Windows.Forms.Label
    $SharePathLabel.Text = "Root"
    #$SharePathLabel.TextAlign = "Left"    
    $SharePathLabel.location = New-Object System.Drawing.Point(($LeftOffset+15),$TopOffset) # (from left, from top)
    $SharePathLabel.Size = New-Object System.Drawing.Size($labelWidth,20)                   # (width, height)
    if ( $v_DebugMode ) { Write-Host 'creating Shared repo text field' }
    $IndividualPathTextField = New-Object System.Windows.Forms.TextBox
    $IndividualPathTextField.Text = "$Script:v_Root_IndividualRepoFolder"       # start w/INI setting
    $IndividualPathTextField.Multiline = $false 
    $IndividualPathTextField.location = New-Object System.Drawing.Point(($LeftOffset+100),($TopOffset-4)) # (from left, from top)
    $IndividualPathTextField.Size = New-Object System.Drawing.Size($PathFieldLength,$FieldHeight)# (width, height)
    $IndividualPathTextField.ReadOnly = $true
    $IndividualPathTextField.Name = "Individual Paths"
    $IndividualPathTextField.add_MouseHover($ShowHelp)
    
    if ( $v_DebugMode ) { Write-Host 'creating 'Individual' Browse button' }
    $sharedBrowse = New-Object System.Windows.Forms.Button
    $sharedBrowse.Width = 60
    $sharedBrowse.Text = 'Browse'
    $sharedBrowse.Location = New-Object System.Drawing.Point(($LeftOffset+$PathFieldLength+$labelWidth+40),($TopOffset-5))
    $sharedBrowse_Click = {
        $IndividualPathTextField.Text = Browse_IndividualRepos $dataGridView $IndividualPathTextField.Text
    } # $sharedBrowse_Click
    $sharedBrowse.add_Click($sharedBrowse_Click)

    #--------------------------------------------------------------------------
    # create radio button, 'common' label, text entry fields, and Browse button 
    #--------------------------------------------------------------------------
    $CommonRadioButton = New-Object System.Windows.Forms.RadioButton
    $CommonRadioButton.Location = '10,34'
    $CommonRadioButton.Add_Click( {
        $Script:v_CommonRepo = $True
        $find = "^[\$]v_CommonRepo"
        $replace = "`$v_CommonRepo = `$$Script:v_CommonRepo" 
        Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"
        Import_CommonRepo $dataGridView $CommonPathTextField.Text $Script:HPModelsTable
        [System.Windows.MessageBox]::Show($OSMessage,$OSHeader,0)    # 1 = "OKCancel" ; 4 = "YesNo"
    } ) # $CommonRadioButton.Add_Click()

    if ( $v_DebugMode ) { Write-Host 'creating Common repo field label' }
    $CommonPathLabel = New-Object System.Windows.Forms.Label
    $CommonPathLabel.Text = "Common"
    $CommonPathLabel.location = New-Object System.Drawing.Point(($LeftOffset+15),($TopOffset+18)) # (from left, from top)
    $CommonPathLabel.Size = New-Object System.Drawing.Size($labelWidth,20)    # (width, height)
    if ( $v_DebugMode ) { Write-Host 'creating Common repo text field' }
    $CommonPathTextField = New-Object System.Windows.Forms.TextBox
    $CommonPathTextField.Text = "$Script:v_Root_CommonRepoFolder"
    $CommonPathTextField.Multiline = $false 
    $CommonPathTextField.location = New-Object System.Drawing.Point(($LeftOffset+100),($TopOffset+15)) # (from left, from top)
    $CommonPathTextField.Size = New-Object System.Drawing.Size($PathFieldLength,$FieldHeight)             # (width, height)
    $CommonPathTextField.ReadOnly = $true
    $CommonPathTextField.Name = "Common Path"
    $CommonPathTextField.add_MouseHover($ShowHelp)
    #$CommonPathTextField.BorderStyle = 'None'                                # 'none', 'FixedSingle', 'Fixed3D (default)'
    
    if ( $v_DebugMode ) { Write-Host 'creating Common Browse button' }
    $commonBrowse = New-Object System.Windows.Forms.Button
    $commonBrowse.Width = 60
    $commonBrowse.Text = 'Browse'
    $commonBrowse.Location = New-Object System.Drawing.Point(($LeftOffset+$PathFieldLength+$labelWidth+40),($TopOffset+13))
    $commonBrowse_Click = {
        Browse_CommonRepo $dataGridView $CommonPathTextField.Text
    } # $commonBrowse_Click
    $commonBrowse.add_Click($commonBrowse_Click)

    $PathsGroupBox.Controls.AddRange(@($IndividualPathTextField, $SharePathLabel,$IndividualRadioButton, $sharedBrowse))
    $PathsGroupBox.Controls.AddRange(@($CommonPathTextField, $CommonPathLabel, $commonBrowse, $CommonRadioButton))

    $CM_form.Controls.AddRange(@($PathsGroupBox))

    #----------------------------------------------------------------------------------
    # Create Models list Checked Grid box - add 1st checkbox column
    # The ListView control allows columns to be used as fields in a row
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating DataGridView to populate with platforms' }
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

    if ( $v_DebugMode ) {  Write-Host 'creating col 0 checkboxColumn' }
    # add column 0 (0 is 1st column)
    $CheckBoxColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $CheckBoxColumn.width = 28

    [void]$DataGridView.Columns.Add($CheckBoxColumn) 

    #----------------------------------------------------------------------------------
    # Add a CheckBox on header (to 1st col)
    # default the all checkboxes selected/checked
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating Checkbox col 0 header' }
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
    if ( $v_DebugMode ) { Write-Host 'adding SysId, Model columns' }
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
    if ( $v_DebugMode ) { Write-Host 'creating category columns' }
    foreach ( $cat in $Script:v_FilterCategories ) {
        $catFilter = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
        $catFilter.name = $cat
        $catFilter.width = 50
        [void]$DataGridView.Columns.Add($catFilter) 
    }
    #----------------------------------------------------------------------------------
    # add an All 'AddOns' column
    # column 7 (0 is 1st column)
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating AddOns column' }
    $CheckBoxINISoftware = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $CheckBoxINISoftware.Name = 'AddOns' 
    $CheckBoxINISoftware.width = 50
    $CheckBoxINISoftware.ThreeState = $False
    [void]$DataGridView.Columns.Add($CheckBoxINISoftware)
   
    # $CheckBoxesAll.state = Displayed, Resizable, ResizableSet, Selected, Visible

    #----------------------------------------------------------------------------------
    # add a repository path as last column
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating Repo links column' }
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
        
        # Let's see if the cell is a checkmark (type Boolean) or a text cell (which would NOT have a value of $true or $false)
        # columns 0=all, 1=sysId, 2=Name, 8=path (all string types)
        if ( $column -in @(0, 3, 4, 5, 6, 7) ) {      # col 7='AddOns'

            # next seems convoluted, but we seem to get 'unchecked' when user clicks on cell, 
            # ... so we reverse that to 'checked' or the new cell state
            $CellprevState = $dataGridView.rows[$row].Cells[$column].EditedFormattedValue # value 'BEFORE' click action $true/$false
            $CellNewState = !$CellprevState                                              # the 'ACTUAL' state of the cell clicked
            
            # here we know we are dealing with one of the checkmark/selection cells
            switch ( $column ) {
                0 {            
                    # 'Driver','BIOS', 'Firmware', 'Software', then 'AddOns'
                    foreach ( $cat in $Script:v_FilterCategories ) {                 # ... all categories          
                        $datagridview.Rows[$row].Cells[$cat].Value = $CellNewState
                    } # forech ( $cat in $Script:v_FilterCategories )
                    $dataGridView.Rows[$row].Cells['AddOns'].Value = $CellNewState  
                    if ( -not $CellNewState ) {
                        $datagridview.Rows[$row].Cells[$datagridview.columnCount-1].Value = '' # ... and reset the repo path field
                    }
                } # 0
                Default {   # here to deal with clicking on a category or 'AddOns' cell 

                    # if we selected the 'AddOns' let's make sure we set this as default by tapping a new file
                    # ... with platform ID as the name... 
                    # it seems as if PS can't distinguish .value from checked or not checked... always the same
                    if ( $datagridview.Rows[$row].Cells['AddOns'].State -eq 'Selected' ) {
                        if ( $Script:v_CommonRepo ) { $lRepoPath = $CommonPathTextField.Text } else { $lRepoPath = $IndividualPathTextField.Text }
                        Manage_AddOnFlag_File $lRepoPath $datagridview[1,$row].value $datagridview[2,$row].value $Null $CellNewState $Script:v_CommonRepo
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
                        } # foreach ( $cat in $Script:v_FilterCategories )
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
    if ( $v_DebugMode ) { Write-Host 'creating Models GroupBox' }

    $CMModelsGroupBox = New-Object System.Windows.Forms.GroupBox
    $CMModelsGroupBox.location = New-Object System.Drawing.Point($LeftOffset,($TopOffset+50))     # (from left, from top)
    $CMModelsGroupBox.Size = New-Object System.Drawing.Size(($ListViewWidth+20),($ListViewHeight+30))       # (width, height)
    $CMModelsGroupBox.text = "HP Models / Repository Category Filters"

    $CMModelsGroupBox.Controls.AddRange(@($dataGridView))

    $CM_form.Controls.AddRange(@($CMModelsGroupBox))

    #----------------------------------------------------------------------------------
    # Add a Use List button
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating a Refresh List button' }
    $RefreshGridButton = New-Object System.Windows.Forms.Button
    $RefreshGridButton.Location = New-Object System.Drawing.Point($LeftOffset,($FormHeight-447))    # (from left, from top)
    $RefreshGridButton.Text = 'Use INI List'
    $RefreshGridButton.Name = 'Use INI List'
    $RefreshGridButton.AutoSize=$true
    $RefreshGridButton.add_MouseHover($ShowHelp)

    $RefreshGridButton_Click={
        #CMTraceLog -Message 'Pupulating HP Models List from INI file $HPModels' -Type $TypeNorm
        Empty_Grid $dataGridView
        Populate_Grid_from_INI $dataGridView $Script:HPModelsTable
    } # $RefreshGridButton_Click={

    $RefreshGridButton.add_Click($RefreshGridButton_Click)

    $CM_form.Controls.Add($RefreshGridButton)
  
    #----------------------------------------------------------------------------------
    # Add a list filters button
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating a list filters button' }
    $ListFiltersdButton = New-Object System.Windows.Forms.Button
    $ListFiltersdButton.Location = New-Object System.Drawing.Point(($LeftOffset+93),($FormHeight-447))    # (from left, from top)
    $ListFiltersdButton.Width = 80
    $ListFiltersdButton.Height = 35
    $ListFiltersdButton.Text = 'Show Filters'
    $ListFiltersdButton.Name = 'Show Filters'
    #$ListFiltersdButton.AutoSize=$true

    $ListFiltersdButton_Click={
        CMTraceLog -Message "Listing filters (calling List_Filters())" -Type $TypeNorm
        List_Filters
    } # $ListFiltersdButton_Click={

    $ListFiltersdButton.add_Click($ListFiltersdButton_Click)

    $CM_form.Controls.Add($ListFiltersdButton)

    #----------------------------------------------------------------------------------
    # Add a Add Model button
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating a Add Model button' }
    $AddModelButton = New-Object System.Windows.Forms.Button
    $AddModelButton.Width = 80
    $AddModelButton.Height = 35
    $AddModelButton.Location = New-Object System.Drawing.Point(($LeftOffset+180),($FormHeight-447))    # (from left, from top)
    $AddModelButton.Text = 'Add Model'
    $AddModelButton.Name = 'Add Model'

    $AddModelButton.add_Click( { 
        if ( $Script:v_CommonRepo ) {
            $lTempPath = $CommonPathTextField.Text } else { $lTempPath = $IndividualPathTextField.Text }

        if ( $null -ne ($lModelArray = AddPlatformForm $DataGridView $Script:v_Softpaqs $lTempPath $Script:v_CommonRepo ) ) {
            $Script:s_ModelAdded = $True
            CMTraceLog -Message "$($lModelArray.ProdCode):$($lModelArray.Name) added to list" -Type $TypeNorm
        } 
    } ) # $AddModelButton.add_Click(

    $CM_form.Controls.Add($AddModelButton)

    #----------------------------------------------------------------------------------
    # Create New Log checkbox
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating New Log checkbox' }
    $NewLogCheckbox = New-Object System.Windows.Forms.CheckBox
    $NewLogCheckbox.Text = 'Start New Log'
    $NewLogCheckbox.Autosize = $true
    $NewLogCheckbox.Location = New-Object System.Drawing.Point(($LeftOffset+280),($FormHeight-450))

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
    # Add a log file field
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating log file field' }

    $LogPathField = New-Object System.Windows.Forms.TextBox
    $LogPathField.Text = "$Script:v_LogFile"
    $LogPathField.Multiline = $false 
    $LogPathField.location = New-Object System.Drawing.Point(($LeftOffset+278),($FormHeight-428)) # (from left, from top)
    $LogPathField.Size = New-Object System.Drawing.Size(340,$FieldHeight)                      # (width, height)
    $LogPathField.ReadOnly = $true
    $LogPathField.Name = "LogPath"
    #$LogPathField.BorderStyle = 'none'                                                  # 'none', 'FixedSingle', 'Fixed3D (default)'
    # next, move cursor to end of text in field, to see the log file name
    $LogPathField.Select($LogPathField.Text.Length,0)
    $LogPathField.ScrollToCaret()

    $CM_form.Controls.AddRange(@($LogPathField ))

    #----------------------------------------------------------------------------------
    # Create Sync button
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating Sync button' }
    $buttonSync = New-Object System.Windows.Forms.Button
    $buttonSync.Width = 90
    $buttonSync.Height = 35
    $buttonSync.Text = 'Sync Repos'
    $buttonSync.Location = New-Object System.Drawing.Point(($FormWidth-130),($FormHeight-447))
    $buttonSync.Name = 'Sync'
    $buttonSync.add_MouseHover($ShowHelp)

    $buttonSync.add_click( {

        if ( ($OSVERComboBox.SelectedItem -in $Script:v_OSVALID10) -or ($OSVERComboBox.SelectedItem -in $Script:v_OSVALID11) ) {
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
                    if ( $v_DebugMode ) { CMTraceLog -Message 'Script connected to CM' -Type $TypeDebug }
                }
            } # if ( $updateCMCheckbox.checked )

            if ( $Script:v_CommonRepo ) {
                Sync_Common_Repository $dataGridView $CommonPathTextField.Text $lCheckedListArray #$Script:s_ModelAdded
                if ( $Script:s_ModelAdded ) { Update_INIModelsList $dataGridView $CommonPathTextField $True }   # $True = head of common repository
            } else {
                sync_individual_repositories $dataGridView $IndividualPathTextField.Text $lCheckedListArray $Script:s_ModelAdded
                if ( $Script:s_ModelAdded ) { Update_INIModelsList $dataGridView $IndividualPathTextField.Text $False }   # $False = head of individual repos
            }
            $Script:s_ModelAdded = $False        # reset as the previous Sync also updated INI file
        } else {
            CMTraceLog -Message 'OS Version not supported' -Type $TypeError
        } # if ( $OSVERComboBox.SelectedItem -in $Script:v_OSVALID ... )

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
    if ( $v_DebugMode ) { Write-Host 'creating Repo Checkbox' }
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
    if ( $v_DebugMode ) { Write-Host 'creating DP Distro update Checkbox' }
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
    if ( $v_DebugMode ) { Write-Host 'creating update HPIA button' }
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
            $HPIAPathField.Text = "$v_HPIACMPackage - $v_HPIAPath"
            CM_HPIAPackage $v_HPIACMPackage $v_HPIAPath $v_HPIAVersion
        }
    } # $HPIAPathButton_Click = 

    $HPIAPathButton.add_Click($HPIAPathButton_Click)

    $HPIAPathField = New-Object System.Windows.Forms.TextBox
    $HPIAPathField.Text = "$v_HPIACMPackage - $v_HPIAPath"
    $HPIAPathField.Multiline = $false 
    $HPIAPathField.location = New-Object System.Drawing.Point(($LeftOffset+120),($TopOffset-5)) # (from left, from top)
    $HPIAPathField.Size = New-Object System.Drawing.Size(320,$FieldHeight)                      # (width, height)
    $HPIAPathField.ReadOnly = $true
    $HPIAPathField.Name = "v_HPIAPath"
    $HPIAPathField.BorderStyle = 'none'                                                  # 'none', 'FixedSingle', 'Fixed3D (default)'

    #----------------------------------------------------------------------------------
    # Create HPIA Browse button
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating HPIA Browse button' }
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
                $HPIAPathField.Text = "$v_HPIACMPackage - $v_HPIAPath"
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
    if ( $v_DebugMode ) { Write-Host 'creating RichTextBox' }
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
    if ( $v_DebugMode ) { Write-Host 'creating a clear textbox checkmark' }
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
    $DebugCheckBox.checked = $Script:v_DebugMode
    $DebugCheckBox.add_click( {
            if ( $DebugCheckBox.checked ) {
                $Script:v_DebugMode = $true
            } else {
                $Script:v_DebugMode = $false
            }
        }
    ) # $DebugCheckBox.add_click

    $CM_form.Controls.Add($DebugCheckBox)                    # removed CM Connect Button

    #----------------------------------------------------------------------------------
    # Add TextBox larger and smaller Font buttons
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating a TextBox smaller Font button' }
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
    if ( $v_DebugMode ) { Write-Host 'creating a TextBox larger Font button' }
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
    if ( $v_DebugMode ) { Write-Host 'creating Done/Exit Button' }
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

    #----------------------------------------------------------------------------------
    # set up the default paths for the repository based on the INI $Script:v_CommonRepo value 
    # ... find variable as long as is at start of a line (otherwise could be a comment)
    #----------------------------------------------------------------------------------
    $find = "^[\$]v_CommonRepo"
    (Get-Content $IniFIleFullPath) | Foreach-Object { 
        if ($_ -match $find) {               # we found the variable
            if ( $_ -match '\$true' ) { 
                $lCommonPath = $CommonPathTextField.Text
                if ( ([string]::IsNullOrEmpty($lCommonPath)) -or (-not (Test-Path $lCommonPath)) ) {
                    Write-Host "Common Repository Path from INI file not found: $($lCommonPath) - Will Create" -ForegroundColor Red
                    Init_Repository $lCommonPath $True                 # $True = make it a HPIA repo
                }   
                Import_CommonRepo $dataGridView $lCommonPath $Script:HPModelsTable
                $CommonRadioButton.Checked = $true                     # set the visual default from the INI setting
                $CommonPathTextField.BackColor = $BackgroundColor
            } else { 
                if ( [string]::IsNullOrEmpty($IndividualPathTextField.Text) ) {
                    Write-Host "Individual Repository field is empty" -ForegroundColor Red
                } else {
                    Import_IndividualRepos $dataGridView $IndividualPathTextField.Text $Script:HPModelsTable
                    $IndividualRadioButton.Checked = $true 
                    $IndividualPathTextField.BackColor = $BackgroundColor
                }
            } # else if ( $_ -match '\$true' )
        } # if ($_ -match $find)
    } # Foreach-Object

    #----------------------------------------------------------------------------------
    # Finally, show the dialog on screen
    #----------------------------------------------------------------------------------

    if ( $v_DebugMode ) { Write-Host 'calling ShowDialog' }
    
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
    MainForm                            # Create the GUI and take over all actions
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
            $Script:v_CommonRepo = $true } else { $Script:v_CommonRepo = $false }  } 
    if ( $PSBoundParameters.Keys.Contains('Products') ) { "-Products: $($Products)" }
    if ( $PSBoundParameters.Keys.Contains('ListFilters') ) { list_filters $Script:v_CommonRepo }
    if ( $PSBoundParameters.Keys.Contains('NoIniSw') ) { '-NoIniSw' }
    if ( $PSBoundParameters.Keys.Contains('showActivityLog') ) { $showActivityLog = $true } 
    if ( $PSBoundParameters.Keys.Contains('Sync') ) { sync_repos $dataGridView $Script:v_CommonRepo }

} # if ( $MyInvocation.BoundParameters.count -gt 0)

########################################################################################