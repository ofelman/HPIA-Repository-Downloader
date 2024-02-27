<#
    HP Image Assistant Softpaq Repository Downloader
    by Dan Felman/HP Technical Consultant
        
        Loop Code: The 'HPModelsTable' loop code, is in a separate INI.ps1 file 
            ... created by Gary Blok's (@gwblok) post on garytown.com.
            ... https://garytown.com/create-hp-bios-repository-using-powershell
 
        Logging: The Log function based on code by Ryan Ephgrave (@ephingposh)
            ... https://www.ephingadmin.com/powershell-cmtrace-log-function/
        Version informtion in Release_Notes.txt

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
    [Switch]$Sync
) # param

$ScriptVersion = "2.04.02 (Feb 23, 2024)"

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
if ( -not (Get-Variable -Name v_OPSYS -ErrorAction SilentlyContinue) ) {
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
Set-Variable s_AddSoftware -Option Constant -Value '.ADDSOFTWARE'
#$s_AddSoftware = '.ADDSOFTWARE'                      # sub-folders where named downloaded Softpaqs will reside
Set-Variable s_HPIAActivityLog -Option Constant -Value 'activity.log'
#$s_HPIAActivityLog = 'activity.log'                  # name of HPIA activity log file
$s_AddAccesories = $false # FUTURE USE #

'My Public IP: '+(Invoke-WebRequest ifconfig.me/ip).Content.Trim() | Out-Host

#--------------------------------------------------------------------------------------
# error codes for color coding, etc.
Set-Variable TypeError -Option Constant -Value -1
Set-Variable TypeNorm -Option Constant -Value 1
Set-Variable TypeWarn -Option Constant -Value 2
Set-Variable TypeDebug -Option Constant -Value 4
Set-Variable TypeSuccess -Option Constant -Value 5
Set-Variable TypeNoNewline -Option Constant -Value 10

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

    # report CMSL version in use
    $hpcmsl = get-module -listavailable | where { $_.name -like 'HPCMSL' }
    $HPCMSLVer = [string]$hpcmsl.Version.Major+'.'+[string]$hpcmsl.Version.Minor+'.'+[string]$hpcmsl.Version.Build
    CMTraceLog -Message "Using CMSL version: $HPCMSLVer" -Type $TypError    

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
        $tc_CMInstall = Split-Path $env:SMS_ADMIN_UI_PATH
        Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1)

        #--------------------------------------------------------------------------------------
        # by now we know the script is running on a CM server and the PS module is loaded
        # so let's get the CMSite content info
    
        Try {
            $Script:SiteCode = (Get-PSDrive -PSProvider CMSite).Name               # assume CM PS modules loaded at this time

            if (Test-Path $tc_CMInstall) {        
                try { Test-Connection -ComputerName "$FileServerName" -Quiet
                    CMTraceLog -Message " ...Connected" -Type $TypeSuccess 
                    $boolConnectionRet = $True
                    $CMGroupAll.Text = 'SCCM - Connected'
                }
                catch {
	                CMTraceLog -Message "Not Connected to File Server, Exiting" -Type $TypeError 
                }
            } else {
                CMTraceLog -Message "CM Installation path NOT FOUND: '$tcCMInstall'" -Type $TypeError 
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
                              $false: do not initialize (used only as root of individual repository folders)

#>
Function init_repository {
    [CmdletBinding()]
	param( $pRepoFolder,
            $pInitialize )
    
    if ( $v_DebugMode ) { CMTraceLog -Message "> init_repository" -Type $TypeNorm }
    $ir_CurrentLoc = Get-Location


    # Make sure folder exists, or create it
    if ( -not (Test-Path $pRepoFolder) ) {    
        Try {
            $ir_ret = New-Item -Path $pRepoFolder -ItemType directory
            CMTraceLog -Message "... repository path was not found, created: $pRepoFolder" -Type $TypeNorm
        } Catch {
            CMTraceLog -Message "... problem: $($_)" -Type $TypeError
        }
    } # if ( -not (Test-Path $pRepoFolder) )

    switch ( $pInitialize ) {
        $true {
            if ( -not (test-path "$pRepoFolder\.Repository") ) {
                Set-Location $pRepoFolder
                $initOut = (Initialize-Repository) 6>&1
                if ( $v_DebugMode ) { CMTraceLog -Message  "... repository Initialized - $($Initout)"  -Type $TypeNorm }
                Set-RepositoryConfiguration -setting OfflineCacheMode -cachevalue Enable 6>&1   
                Set-RepositoryConfiguration -setting RepositoryReport -Format csv 6>&1  
                if ( $v_DebugMode ) { CMTraceLog -Message  "... configuring $($pRepoFolder) for HP Image Assistant" -Type $TypeNorm }
            } # if ( -not (test-path "$pRepoFolder\.Repository") )
            # next, intialize .ADDSOFTWARE folder for holding named softpaqs
            $ir_AddSoftpaqsFolder = "$($pRepoFolder)\$($s_AddSoftware)"
            if ( -not (Test-Path $ir_AddSoftpaqsFolder) ) {
                if ( $v_DebugMode ) { CMTraceLog -Message "... creating Add-on Softpaqs Folder $ir_AddSoftpaqsFolder" -Type $TypeNorm }
                $lret = New-Item -Path $ir_AddSoftpaqsFolder -ItemType directory
            } # if ( !(Test-Path $ir_AddSoftpaqsFolder) )
        } # $true
        $false {
            if ( $v_DebugMode ) { CMTraceLog -Message  "... Folder $($pRepoFolder) available" -Type $TypeNorm }
        } # $false
    } # switch ( $pInitialize )

    if ( $v_DebugMode ) { CMTraceLog -Message "< init_repository" -Type $TypeNorm }
    Set-Location $ir_CurrentLoc
} # Function init_repository

#=====================================================================================
<#
    Function Download_Softpaq
        This function will download a Softpaq, and its CVA and HTML associated files
        if Softpaq previously downloaded, will not redownload, but the CVA dn HTML files are
        Args:
            $pSoftpaq:          softpaq to download
            $pAddOnFolderPath:  folder to download files to
#>
Function Download_Softpaq {
    param( $pSoftpaq,
            $pAddOnFolderPath )
    
    $da_CurrLocation = Get-Location

    Set-Location $pAddOnFolderPath

    $da_SoftpaqExePath = $pAddOnFolderPath+'\'+$pSoftpaq.id+'.exe'
    if ( Test-Path $da_SoftpaqExePath ) {
        CMTraceLog -Message "`t$($pSoftpaq.id) already downloaded - $($pSoftpaq.Name)" -Type $TypeWarn
    } else {                                        
        CMTraceLog -Message "`tdownloading $($pSoftpaq.id)/$($pSoftpaq.Name)" -Type $TypeNoNewline
        Try {
            $ret = (Get-Softpaq $pSoftpaq.Id) 6>&1
            CMTraceLog -Message "... done" -Type $TypeNorm
        } Catch {
            CMTraceLog -Message "... failed to download: $($_)" -Type $TypeError
        }          
    } # else if ( Test-Path $da_SoftpaqExePath ) {
    
    # download corresponding CVA file
    CMTraceLog -Message "`tdownloading $($pSoftpaq.id) CVA file" -Type $TypeNoNewline
    Try {
        $ret = (Get-SoftpaqMetadataFile $pSoftpaq.Id -Overwrite 'Yes') 6>&1   # ALWAYS update the CVA file, even if previously downloaded
        CMTraceLog -Message "... done" -Type $TypeNorm    
    } Catch {
        CMTraceLog -Message "... failed to download: $($_)" -Type $TypeError
    }
    
    # download corresponding HTML file
    CMTraceLog -Message "`tdownloading $($pSoftpaq.id) HTML file" -Type $TypeNoNewline
    Try {
        $da_SoftpaqHtml = $pAddOnFolderPath+'\'+$pSoftpaq.id+'.html' # where to download to
        $ret = Invoke-WebRequest -UseBasicParsing -Uri $pSoftpaq.ReleaseNotes -OutFile "$da_SoftpaqHtml"
        CMTraceLog -Message "... done" -Type $TypeNorm
    } Catch {
        CMTraceLog -Message "... failed to download" -Type $TypeError
    }    
    
    Set-Location $da_CurrLocation
} # Function Download_Softpaq

Function Get_Softpaqs {
[CmdletBinding()] 
    param( $pAddOnFlagFileFullPath )

    $gs_ProdCode = Split-Path $pAddOnFlagFileFullPath -leaf
    $gs_AddSoftpaqsFolder = Split-Path $pAddOnFlagFileFullPath -Parent

    [array]$gs_AddOnsList = Get-Content $pAddOnFlagFileFullPath

    if ( $gs_AddOnsList.count -ge 1 ) {
        CMTraceLog -Message "... platform $($gs_ProdCode): checking AddOns flag file"
        if ( $v_DebugMode ) { CMTraceLog -Message 'calling Get-SoftpaqList():'+$gs_ProdCode -Type $TypeNorm }
        Try {
            $gs_SoftpaqList = (Get-SoftpaqList -platform $gs_ProdCode -os $Script:v_OS -osver $Script:v_OSVER -characteristic SSM) 6>&1  
            # check every reference file Softpaq for a match in the flg file list
            ForEach ( $gs_iEntry in $gs_AddOnsList ) {
                $gs_EntryFound = $False
                CMTraceLog -Message "... Checking AddOn Softpaq by name: $($gs_iEntry)" -Type $TypeNorm
                CMTraceLog -Message "... # entries in ref file: $($gs_iEntry.Count)" -Type $TypeNorm
                ForEach ( $gs_iSoftpaq in $gs_SoftpaqList ) {
                    if ( [string]$gs_iSoftpaq.Name -match $gs_iEntry ) {
                        $gs_EntryFound = $True
                        CMTraceLog -Message "... calling Download_Softpaq(): $($gs_iEntry)"
                        Download_Softpaq $gs_iSoftpaq $gs_AddSoftpaqsFolder
                    } # if ( [string]$gs_iSoftpaq.Name -match $gs_iEntry )
                } # ForEach ( $gs_iSoftpaq in $gs_SoftpaqList )
                if ( -not $gs_EntryFound ) {
                    CMTraceLog -Message  "... '$($gs_iEntry)': Softpaq not found for this platform and OS version"  -Type $TypeWarn
                } # if ( -not $gs_EntryFound )
            } # ForEach ( $lEntry in $gs_AddOnsList )
        } Catch {
            CMTraceLog -Message "... $($gs_ProdCode): Error retrieving Reference file" -Type $TypeError
        }                        
    } else {
        CMTraceLog -Message $gs_ProdCode': Flag file found but empty '
    } # else if ( -not ([String]::IsNullOrWhiteSpace($lAddOns)) )

} # Function Get_Softpaqs

#=====================================================================================
<#
    Function Download_AddOns
        designed to find named Softpaqs from $HPModelsTable and download those into their own repo folder
        Requires: pGrid - entries from the UI grid devices
                  pFolder - repository folder to check out
                  pCommonRepoFlag - check for using Common repo for all devices
#>
Function Download_AddOns {
[CmdletBinding()] 
    param( $pGrid,
            $pFolder,
            $pCommonRepoFlag )

    if ( $script:noIniSw ) { return }    # $script:noIniSw = runstring option, if executed from command line

    $da_CurrentLoc = Get-Location

    $da_AddSoftpaqsFolder = "$($pFolder)\$($s_AddSoftware)"   # get path of .ADDSOFTWARE subfolder
    Set-Location $da_AddSoftpaqsFolder

    switch ( $pCommonRepoFlag ) {
        #--------------------------------------------------------------------------------
        # Download AddOns for every device with flag file content 
        #--------------------------------------------------------------------------------
        $true { 
            #--------------------------------------------------------------------------------
            # There are potentially multiple models being held in this repository
            # ... so find all Platform AddOns flag files for content defined by user
            #--------------------------------------------------------------------------------
            for ( $iRow = 0; $iRow -lt $pGrid.RowCount; $iRow++ ) {
        
                $da_ProdCode = $pGrid[1,$iRow].Value                      # column 1 = Model/Prod ID
                $da_ModelName = $pGrid[2,$iRow].Value                     # column 2 = Model name
                $lAddOnsFlag = $pGrid.Rows[$iRow].Cells['AddOns'].Value # column 7 = 'AddOns' checkmark
                $lProdIDFlagFile = $da_AddSoftpaqsFolder+'\'+$da_ProdCode

                # if user checked the addOns column... and the flag file is there...
                if ( $lAddOnsFlag -and (Test-Path $lProdIDFlagFile) ) {                    

                    if ( $v_DebugMode ) { CMTraceLog -Message 'Flag file found:'+$lAddOnsFlagFile -Type $TypeWarn }
                    Get_Softpaqs $lAddOnsFlagFile 
<#
                    [array]$da_AddOnsList = Get-Content $lProdIDFlagFile

                    if ( $da_AddOnsList.count -ge 1 ) {
                        CMTraceLog -Message "... platform $($da_ProdCode): checking AddOns flag file"
                        if ( $v_DebugMode ) { CMTraceLog -Message 'calling Get-SoftpaqList():'+$da_ProdCode -Type $TypeNorm }
                        Try {
                            $da_SoftpaqList = (Get-SoftpaqList -platform $da_ProdCode -os $Script:v_OS -osver $Script:v_OSVER -characteristic SSM) 6>&1  
                            # check every reference file Softpaq for a match in the flg file list
                            ForEach ( $da_iEntry in $da_AddOnsList ) {
                                $da_EntryFound = $False
                                CMTraceLog -Message "... Checking AddOn Softpaq by name: $($da_iEntry)" -Type $TypeNorm
                                CMTraceLog -Message "... # entries in ref file: $($da_iEntry.Count)" -Type $TypeNorm
                                ForEach ( $iSoftpaq in $da_SoftpaqList ) {
                                    if ( [string]$iSoftpaq.Name -match $da_iEntry ) {
                                        $da_EntryFound = $True
                                        CMTraceLog -Message "... calling Download_Softpaq(): $($da_iEntry)"
                                        Download_Softpaq $iSoftpaq $da_AddSoftpaqsFolder
                                    } # if ( [string]$iSoftpaq.Name -match $da_iEntry )
                                } # ForEach ( $iSoftpaq in $da_SoftpaqList )
                                if ( -not $da_EntryFound ) {
                                    CMTraceLog -Message  "... '$($da_iEntry)': Softpaq not found for this platform and OS version"  -Type $TypeWarn
                                } # if ( -not $da_EntryFound )
                            } # ForEach ( $lEntry in $da_AddOnsList )
                        } Catch {
                            CMTraceLog -Message "... $($da_ProdCode): Error retrieving Reference file" -Type $TypeError
                        }                        
                    } else {
                        CMTraceLog -Message $da_ProdCode': Flag file found but empty '
                    } # else if ( -not ([String]::IsNullOrWhiteSpace($lAddOns)) )
#>
                } else {
                    CMTraceLog -Message '... '$da_ProdCode': AddOns cell not checked, will not attempt to download'
                } # else if (Test-Path $lAddOnsFlagFile)

            } # for ( $i = 0; $i -lt $pGrid.RowCount; $i++ )
        } # $true

        $false {    # Download AddOns for individual models ( $pCommonRepoFlag = $false )
            #--------------------------------------------------------------------------------
            # Search grid for SysID, so we can find the AddOns flag file, and check it
            #--------------------------------------------------------------------------------
            for ( $i = 0; $i -lt $pGrid.RowCount; $i++ ) {
                $da_SystemID = $pGrid[1,$i].Value                           # column 1 has the Model/Prod ID
                $lSystemName = $pGrid[2,$i].Value                           # column 2 has the Model name
                $lFoldername = split-path $pFolder -Leaf

                if ( $lFoldername -match $lSystemName ) {                   # we found the model, so get the Prod ID
                    $da_AddOnsValue = $pGrid.Rows[$i].Cells['AddOns'].Value   # column 7 has the AddOns checkmark
                    $lProdIDFlagFile = $da_AddSoftpaqsFolder+'\'+$da_SystemID

                    if ( $da_AddOnsValue -and (Test-Path $lProdIDFlagFile) ) {                        

                        if ( $v_DebugMode ) { CMTraceLog -Message 'Flag file found:'+$lProdIDFlagFile -Type $TypeWarn }
                        Get_Softpaqs $lProdIDFlagFile
<#              
                        [array]$da_AddOnsList = Get-Content $lProdIDFlagFile
                        if ( $da_AddOnsList.count -ge 1 ) {
                            CMTraceLog -Message "  ... AddOns Softpaq entries: $($da_AddOnsList)" -Type $TypeNorm
                            Try {
                                $da_SoftpaqList = (Get-SoftpaqList -platform $da_SystemID -os $Script:v_OS -osver $Script:v_OSVER -characteristic SSM) 6>&1  
                                ForEach ( $da_iEntry in $da_AddOnsList ) {     # search available Softpaq for a match
                                    $da_EntryFound = $False
                                    $da_SoftpaqList | foreach { 
                                        if ( $_.Name -match $da_iEntry ) {
                                            $da_EntryFound = $true
                                            CMTraceLog -Message "  ... Downloading AddOn Softpaq: $($_.Name)"
                                            Download_Softpaq $_ $da_AddSoftpaqsFolder
                                        } # if ( $_.Name -match $da_iEntry )
                                    } # $da_SoftpaqList | foreach
                                    if ( -not $da_EntryFound ) {
                                        CMTraceLog -Message "  ... AddOn Softpaq by name not found: $($da_iEntry)"
                                    }
                                } # ForEach ( $lEntry in $da_AddOnsList )
                            } Catch {
                                CMTraceLog -Message "  ... $($da_SystemID): Error retrieving Reference file" -Type $TypeError
                            }
                        } else {
                            CMTraceLog -Message '  ... '$da_SystemID': AddOn flag file found but empty'
                        } # else if ( $da_AddOnsList.count -ge 1 )
#>                        
                    } # if ( $da_AddOnsValue -and (Test-Path $lProdIDFlagFile) )
                    
                } # if ( $lFoldername -match $lSystemName )
            } # for ( $i = 0; $i -lt $pGrid.RowCount; $i++ )
        } # false

    } #  switch ( $pCommonRepoFlag )
    
    Set-Location $da_CurrentLoc

} # Function Download_AddOns

#=====================================================================================
<#
    Function Restore_AddOns
        This function will restore softpaqs from .ADDSOFTWARE to root

    expects Parameter 
        - Repository folder to sync
#>
Function Restore_AddOns {
    [CmdletBinding()]
	param( $pFolder, 
            $pAddSoftpaqs )        # $True = restore Softpaqs to repo from .ADDSOTWARE folder
    
        if ( $pAddSoftpaqs ) {
            CMTraceLog -Message "... Restoring named softpaqs/cva files to Repository"
            $ras_source = "$($pFolder)\$($s_AddSoftware)\*.*"
            Copy-Item -Path $ras_source -Filter *.cva -Destination $pFolder
            Copy-Item -Path $ras_source -Filter *.exe -Destination $pFolder
            Copy-Item -Path $ras_source -Filter *.html -Destination $pFolder
            CMTraceLog -Message "... Restoring named softpaqs/cva files completed"
        } # if ( -not $noIniSw )

} # Restore_AddOns

#=====================================================================================
<#
    Function Sync_and_Cleanup_Repository
        This function will run a Sync and a Cleanup commands from HPCMSL

    Parameter 
        - pGrid
        - Repository folder to sync
        - Flag : True if this is a common repo, False if an individual repo
#>
Function Sync_and_Cleanup_Repository {
    [CmdletBinding()]
	param( $pGrid, $pFolder, $pCommonFlag )

    $scr_CurrentLoc = Get-Location
    CMTraceLog -Message  "    >> Sync_and_Cleanup_Repository - '$pFolder' - please wait !!!" -Type $TypeNorm

    if ( Test-Path $pFolder ) {
        #--------------------------------------------------------------------------------
        # update repository softpaqs with sync command and then cleanup
        #--------------------------------------------------------------------------------
        Set-Location -Path $pFolder
        
        Try {
            CMTraceLog -Message  '...   calling CMSL Invoke-RepositorySync' -Type $TypeNorm
            $lres = invoke-repositorysync 6>&1
                
            # find what sync'd from the CMSL log file for this run
            try {
                CMTraceLog -Message  '...   retrieving repository filters' -Type $TypeNorm
                #$lRepoFilters = (Get-RepositoryInfo).Filters  # see if any filters are used to continue

                Get_ActivityLogEntries $pFolder  # get sync'd info from HPIA activity log file
                $sc_contentsFile = "$($pFolder)\.repository\contents.csv"
                if ( Test-Path $sc_contentsFile ) {
                    $lContentsHASH = (Get-FileHash -Path "$($pFolder)\.repository\contents.csv" -Algorithm SHA256).Hash
                    if ( $v_DebugMode ) { CMTraceLog -Message "... SHA256 Hash of 'contents.csv': $($lContentsHASH)" -Type $TypeWarn }
                }              
                #--------------------------------------------------------------------------------
                CMTraceLog -Message  '...   calling CMSL invoke-RepositoryCleanup' -Type $TypeNorm
                invoke-RepositoryCleanup 6>&1

                CMTraceLog -Message  '...   calling Download_AddOns()' -Type $TypeNorm
                Download_AddOns $pGrid $pFolder $pCommonFlag

                CMTraceLog -Message  '...   calling Restore_Softpaqs()' -Type $TypeNorm            
                Restore_AddOns $pFolder (-not $script:noIniSw)
                #--------------------------------------------------------------------------------
                # see if Cleanup modified the contents.csv file 
                # - seems like (up to at least 1.6.3) RepositoryCleanup does not Modify 'contents.csv'
                if ( Test-Path $sc_contentsFile ) {
                    $lContentsHASH = Get_SyncdContents $pFolder       # Sync command creates 'contents.csv'
                    CMTraceLog -Message "...   MD5 Hash of 'contents.csv': $($lContentsHASH)" -Type $TypeWarn
                }                
            } catch {
                $lContentsHASH = $null
            }
        } Catch {
            CMTraceLog -Message  "...   error w/Sync $_" -Type $TypeError
        } # Catch Try
        CMTraceLog -Message  '    << Sync_and_Cleanup_Repository - Done' -Type $TypeNorm

    } # if ( Test-Path $pFolder )

    Set-Location $scr_CurrentLoc

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

    $umf_CurrentLoc = Get-Location

    set-location $pModelRepository

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
            CMTraceLog -Message  "... adding filter: -Platform $pModelID -os $Script:v_OS:$Script:v_OSVER -category $($cat) -characteristic ssm" -Type $TypeNoNewline
            $lRes = (Add-RepositoryFilter -platform $pModelID -os $Script:v_OS -osver $Script:v_OSVER -category $cat -characteristic ssm 6>&1)          
            CMTraceLog -Message $lRes -Type $TypeWarn 
        } # if ( $pGrid.Rows[$pRow].Cells[$cat].Value )
    } # foreach ( $cat in $Script:v_FilterCategories )
    #--------------------------------------------------------------------------------
    # update repository path field for this model in the grid (path is the last col)
    #--------------------------------------------------------------------------------
    $pGrid[($pGrid.ColumnCount-1),$pRow].Value = $pModelRepository

    Set-Location -Path $umf_CurrentLoc

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
    Function sync_individualRepositories
        for every selected model, go through every repository by model
            - ensure there is a valid repository
            - remove all filters
            - add filters as selected in UI
            - invoke sync and cleanup function on each folder 
#>
Function sync_individualRepositories {
[CmdletBinding()]
	param( $pGrid,
            $pRepoHeadFolder,                                            
            $pCheckedItemsList,
            $pNewModels )                                 # $True = models added to list)                                      # array of rows selected
    
    CMTraceLog -Message "> sync_individualRepositories - START" -Type $TypeNorm
    if ( $v_DebugMode ) { CMTraceLog -Message "... for OS version: $($Script:v_OS):$($Script:v_OSVER)" -Type $TypeDebug }
    if ( $v_DebugMode ) { CMTraceLog -Message "... Folder: $($pRepoHeadFolder)" -Type $TypeDebug }
    if ( $v_DebugMode ) { CMTraceLog -Message "... checkeditems: $($pCheckedItemsList)" -Type $TypeDebug }

    #--------------------------------------------------------------------------------
    # decide if we work on single repo or individual repos
    #--------------------------------------------------------------------------------     
    init_repository $pRepoHeadFolder $false             # make sure Main repo folder exists, do NOT make it a repository

    #--------------------------------------------------------------------------------
    if ( $v_DebugMode ) { CMTraceLog -Message "... stepping through all selected models" -Type $TypeDebug }

    # go through every selected HP Model in the list
    for ( $i = 0; $i -lt $pGrid.RowCount; $i++ ) {
        if ( $v_DebugMode ) { CMTraceLog -Message "... stepping through grid row "+$i -Type $TypeDebug }
        $lModelId = $pGrid[1,$i].Value                    # column 1 has the Model/Prod ID
        $si_ModelName = $pGrid[2,$i].Value                  # column 2 has the Model name

        # if model entry is checked, we need to create a repository, but ONLY if $Script:v_CommonRepo = $false
        if ( $i -in $pCheckedItemsList ) {

            CMTraceLog -Message "`n... Updating model: $lModelId : $si_ModelName"
            if ( $v_DebugMode ) { CMTraceLog -Message "... stepping through model "+$si_ModelName -Type $TypeDebug }
            $sir_TempRepoFolder = "$($pRepoHeadFolder)\$($lModelId)_$($si_ModelName)"     # this is the repo folder for this model

            if ( -not (Test-Path -Path $sir_TempRepoFolder) ) {
                $sir_TempRepoFolder = "$($pRepoHeadFolder)\$($si_ModelName)"           # this is the repo folder for this model, without SysID in name
            }
            if ( $v_DebugMode ) { CMTraceLog -Message "... stepping through path "+$sir_TempRepoFolder -Type $TypeDebug }
            
            init_repository $sir_TempRepoFolder $true
            #--------------------------------------------------------------------------------
            # rebuild repository filters
            #--------------------------------------------------------------------------------
            Update_Model_Filters $pGrid $sir_TempRepoFolder $lModelId $i
            #--------------------------------------------------------------------------------
            # now sync up and cleanup this repository
            #--------------------------------------------------------------------------------
            Sync_and_Cleanup_Repository $pGrid $sir_TempRepoFolder $False # sync individual repo
            #--------------------------------------------------------------------------------
            # update SCCM Repository package, if user checked that off
            #--------------------------------------------------------------------------------
            if ( $Script:v_UpdateCMPackages ) {
                CM_RepoUpdate $si_ModelName $lModelId $sir_TempRepoFolder
            }
            
        } # if ( $i -in $pCheckedItemsList )

    } # for ( $i = 0; $i -lt $pGrid.RowCount; $i++ )

    if ( $pNewModels ) { Update_INIModelsListFromGrid $pGrid $pCommonRepoFolder }

    CMTraceLog -Message "< sync_individualRepositories DONE" -Type $TypeSuccess

} # Function sync_individualRepositories


#=====================================================================================
<#
    Function Sync_CommonRepository
        for every selected model, 
            - ensure there is a valid repository
            - remove all filters
            - add filters as selected in UI
            - invoke sync and cleanup function on the folder 
#>
Function Sync_CommonRepository {
    [CmdletBinding()]
        param( $pGrid,                                  # the UI grid of platforms
                $pRepoFolder,                           # location of repository
                $pCheckedItemsList )                    # list of rows checked
    
        if  ( ($pCheckedItemsList).count -eq 0 ) { return }
    
        $pCurrentLoc = Get-Location

        CMTraceLog -Message "> Sync_CommonRepository - START" -Type $TypeNorm
        if ( $v_DebugMode ) { CMTraceLog -Message "... for OS version: $($Script:v_OSVER)" -Type $TypeDebug }
    
        init_repository $pRepoFolder $true              # make sure repo folder exists, or create it, initialize
    
        if ( $Script:v_KeepFilters ) {
            CMTraceLog -Message  "... Keeping existing filters in: $($pRepoFolder)" -Type $TypeNorm
        } else {
            CMTraceLog -Message  "... Removing existing filters in: $($pRepoFolder)" -Type $TypeNorm
        } # else if ( $Script:v_KeepFilters )
          
        #-------------------------------------------------------------------------------------------
        # loop through every Model in the grid, and look for selected rows in UI (left cell checked)
        #-------------------------------------------------------------------------------------------
        for ( $iGridRow = 0; $iGridRow -lt $pGrid.RowCount; $iGridRow++ ) {
            
            $lModelSelected = $pGrid[0,$iGridRow].Value             # col 0 checked if any category checked
            $lModelId = $pGrid[1,$iGridRow].Value                   # col 1 has the SysID
            $sr_ModelName = $pGrid[2,$iGridRow].Value               # col 2 has the Model name
            #$lAddOnsFlag = $pGrid.Rows[$iGridRow].Cells['AddOns'].Value # column 7 has the AddOns checkmark
    
            #---------------------------------------------------------------------------------
            # Remove existing filter for this platform, unless the KeepFilters checkbox is set
            #---------------------------------------------------------------------------------
            if ( -not $Script:v_KeepFilters ) {
                if ( ((Get-RepositoryInfo).Filters).count -gt 0 ) {
                    $lres = (Remove-RepositoryFilter -platform $lModelID -yes 6>&1)
                    if ( $v_DebugMode ) { CMTraceLog -Message "... removed filters for: $($lModelID)" -Type $TypeWarn }
                }
            } # if ( $Script:v_KeepFilters )
    
            if ( $lModelSelected ) {    
                CMTraceLog -Message "... Updating model: $lModelId : $sr_ModelName"
                Update_Model_Filters $pGrid $pRepoFolder $lModelId $iGridRow
                #--------------------------------------------------------------------------------
                # update SCCM Repository package, if enabled in the UI 
                #--------------------------------------------------------------------------------
                if ( $Script:v_UpdateCMPackages ) { CM_RepoUpdate $sr_ModelName $lModelId $pRepoFolder }
            } # if ( $lModelSelected )
        } # for ( $iGridRow = 0; $iGridRow -lt $pGrid.RowCount; $iGridRow++ )
    
        #-----------------------------------------------------------------------------------
        # we are done checking every model for filters, so now do a Softpaq Sync and cleanup
        #-----------------------------------------------------------------------------------
        Sync_and_Cleanup_Repository $pGrid $pRepoFolder $True # $True = sync a common repository
    
        #if ( $pNewModels ) { Update_INIModelsListFromGrid $pGrid $pRepoFolder }   # $False = head of individual repos
    
        #--------------------------------------------------------------------------------
        Set-Location -Path $pCurrentLoc
    
        CMTraceLog -Message "< Sync_CommonRepository DONE" -Type $TypeSuccess
    
} # Function Sync_CommonRepository
    
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
        $lf_platforms = $lProdFilters.platform | Get-Unique

        # develop the filter list by product
        foreach ( $iProduct in $lf_platforms ) {
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
                    } # if ( Test-Path $lAddOnsRepoFile )
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

        $gi_ModelId = $pGrid[1,$iRow].Value                               # column 1 has the Model/Prod ID
        $gi_ModelName = $pGrid[2,$iRow].Value                             # column 2 has the Model name 
        $gi_TempRepoFolder = "$($pRepoRoot)\$($gi_ModelId)_$($gi_ModelName)"  # this is the repo folder - SysID in the folder name
        if ( -not (Test-Path -Path $gi_TempRepoFolder) ) {
            $gi_TempRepoFolder = "$($pRepoRoot)\$($gi_ModelName)"           # this is the repo folder for this model, without SysID in name
        }

        if ( Test-Path $gi_TempRepoFolder ) {
            set-location $gi_TempRepoFolder                               # move to location of Repository
                                    
            $pGrid[($dataGridView.ColumnCount-1),$iRow].Value = $gi_TempRepoFolder # set the location of Repository in the grid row
            
            # if we are maintaining AddOns Softpaqs, show it in the checkbox
            $gi_AddOnRepoFile = $gi_TempRepoFolder+'\'+$s_AddSoftware+'\'+$gi_ModelId   
            if ( Test-Path $gi_AddOnRepoFile ) {   
                if ( [String]::IsNullOrWhiteSpace((Get-content $gi_AddOnRepoFile)) ) {
                            $pGrid.Rows[$iRow].Cells['AddOns'].Value = $False
                } else {
                    $pGrid.Rows[$iRow].Cells['AddOns'].Value = $True
                    [array]$lAddOns = Get-Content $gi_AddOnRepoFile
                    $lMsg = "... Additional Softpaqs Enabled for Platform '$gi_ModelId' {$lAddOns}"
                    CMTraceLog -Message $lMsg -Type $TypeWarn 
                }
            } # if ( Test-Path $gi_AddOnRepoFile )

            $gi_ProdFilters = (get-repositoryinfo).Filters
            <#
                platform        : 8715
                operatingSystem : win10:21h1
                category        : BIOS
                releaseType     : *
                characteristic  : ssm
            #>
            $gi_platformList = @()
            foreach ( $iEntry in $gi_ProdFilters ) {
                if ( $v_DebugMode ) { CMTraceLog -Message "... Platform $($gi_ProdFilters.platform) ... $($gi_ProdFilters.operatingSystem) $($gi_ProdFilters.category) $($gi_ProdFilters.characteristic) - @$($gi_TempRepoFolder)" -Type $TypeWarn }
                if ( -not ($iEntry.platform -in $gi_platformList) ) {
                    $gi_platformList += $iEntry.platform
                    if ( $v_DebugMode ) { CMTraceLog -Message "... populating filter categories ''$($gi_ModelName)''" -Type $TypeNorm }
                    foreach ( $cat in  ($gi_ProdFilters.category.split(' ')) ) {
                        $pGrid.Rows[$iRow].Cells[$cat].Value = $true
                    }
                } # else 
            } # foreach ( $iEntry in $gi_ProdFilters )
        } # if ( Test-Path $gi_TempRepoFolder )
    } # for ( $iRow = 0; $iRow -lt $pModelsList.RowCount; $iRow++ ) 

    if ( $v_DebugMode ) { CMTraceLog -Message '< Get_IndividualRepofilters]' -Type $TypeNorm }
    
} # Function Get_IndividualRepofilters

#=====================================================================================
<#
    Function Empty_Grid
        removes (empties out) all rows from the grid
#>
Function Empty_Grid {
    [CmdletBinding()]
	param( $pGrid )

    for ( $i = $pGrid.RowCount; $i -gt 0; $i-- ) {
        $pGrid.Rows.RemoveAt($i-1)
    } # for ( $i = $pGrid.RowCount; $i -gt 0; $i-- )          

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

    CMTraceLog -Message ">> Browse Individual Repositories - Start" -Type $TypeSuccess

    # ask the user for the repository

    $bi_browse = New-Object System.Windows.Forms.FolderBrowserDialog
    $bi_browse.SelectedPath = $pCurrentRepository         
    $bi_browse.Description = "Select a Root Folder for individual repositories"
    $bi_browse.ShowNewFolderButton = $true
                                  
    if ( $bi_browse.ShowDialog() -eq "OK" ) {
        $bi_Repository = $bi_browse.SelectedPath
        Import_RootedRepos $pGrid $bi_browse.SelectedPath
    } else {
        $bi_Repository = $null
    } # else if ( $bi_browse.ShowDialog() -eq "OK" )
    
    $bi_browse.Dispose()

    CMTraceLog -Message "<< Browse Individual Repositories - Done" -Type $TypeSuccess

    Return $bi_Repository

} # Function Browse_IndividualRepos

#=====================================================================================
<#
    Function Import_RootedRepos
        - Populate list of Platforms from Individual Repositories
        - also, update INI file about imported repository models if user agrees
        Paramaters: $pGrid = pointer to model grid in UI
                    $pCurrentRepository = head of individual repository folders
                    
        returns: object w/Folder picked for repository, or $null if user cancelled
                Folder return value contains True/Fail entry [0] + folder name [1]
#>
Function Import_RootedRepos {
    [CmdletBinding()]
	param( $pGrid, $pCurrentRepository )
    CMTraceLog -Message "> Import_RootedRepos START - at ($($pCurrentRepository))" -Type $TypeNorm

    CMTraceLog -Message  "... calling Empty_Grid() to clear list" -Type $TypeNorm
    Empty_Grid $pGrid
    if ( -not (Test-Path $pCurrentRepository) ) {
        init_repository $pCurrentRepository $False  # $false means this is not a HPIA repo, so don't init as such
    }
    Set-Location $pCurrentRepository
    CMTraceLog -Message  "... checking for existing repositories" -Type $TypeNorm
    $ir_Directories = (Get-ChildItem -Path $pCurrentRepository -Directory) | where { $_.BaseName -notmatch "_BK$" }
    # add a row in the DataGrid for every repository (subfolders) at this location
    $ir_ReposFound = $False     # assume no repositories exist here
    foreach ( $iFolder in $ir_Directories ) { 
        $ir_RepoFolder = $pCurrentRepository+'\'+$iFolder            
        Set-Location $ir_RepoFolder
        Try {
            CMTraceLog -Message  "... [Repository Found ($($ir_RepoFolder)) -- adding model(s) to grid]" -Type $TypeNorm
            # obtain platform SysID from filters in the repository
            [array]$ir_RepoPlatforms = (Get-RepositoryInfo).Filters.platform
            $ir_RepoPlatforms | foreach {
                [void]$pGrid.Rows.Add(@( $true, $_, $iFolder ))
            } # $ir_RepoPlatforms | foreach           
            $ir_ReposFound = $True
        } Catch {
            CMTraceLog -Message  "... $($ir_RepoFolder) is NOT a Repository" -Type $TypeNorm
        } # Catch
    } # foreach ( $iFolder in $ir_Directories )
    
    if ( $ir_ReposFound ) {
        CMTraceLog -Message  "... calling Get_IndividualRepofilters() - retrieving Filters"
        Get_IndividualRepofilters $pGrid $pCurrentRepository
        if ( $v_DebugMode ) { CMTraceLog -Message  "... calling Update_INIModelsListFromGrid()" }
        Update_INIModelsListFromGrid $pGrid $pCurrentRepository  # $False = treat as head of individual repositories
    } else {
        CMTraceLog -Message  "... no repositories found"
        #$bi_ask = [System.Windows.MessageBox]::Show('Clear the HPModel list in the INI file?','INI File update','YesNo')
        $bi_ask = AskUser 'Clear the HPModel list in the INI file?' 4 # 4="YesNo"
        if ( $bi_ask -eq 'Yes' ) {
            CMTraceLog -Message  "...calling Update_UIandINISetting() to clear HPModels list "
            Update_UIandINISetting $pCurrentRepository $False              # $False = using individual repositories
        } else {
            CMTraceLog -Message  "... previous HPModels list remain in INI file"
        }
    } # else if ( $ir_ReposFound )

    CMTraceLog -Message '< Import_RootedRepos DONE' -Type $TypeSuccess
} # Function Import_RootedRepos

#=====================================================================================
<#
    Function Browse_CommonRepo
        Browse to find existing or create new repository
        Paramaters: $pGrid = pointer to model grid in UI
                    $pCurrentRepository
                    $pModelsTable
        returns: object w/Folder picked for repository, or $null if user cancelled
                Folder return value contains True/Fail entry [0] + folder name [1]
#>
Function Browse_CommonRepo {
    [CmdletBinding()]
	param( $pGrid, $pCurrentRepository, $pModelsTable )                                

    CMTraceLog -Message '> Browse_CommonRepo START' -Type $TypeNorm
    # let's find the repository to import or use the current

    $bc_browse = New-Object System.Windows.Forms.FolderBrowserDialog
    $bc_browse.SelectedPath = $pCurrentRepository        # start with the repo listed in the INI file
    $bc_browse.Description = "Select an HPIA Common/Shared Repository"
    $bc_browse.ShowNewFolderButton = $true

    if ( $bc_browse.ShowDialog() -eq "OK" ) {
        $bc_Repository = $bc_browse.SelectedPath
        Import_Repository $pGrid $bc_Repository $pModelsTable    
        CMTraceLog -Message "< Browse Common Repository Done ($($bc_Repository))" -Type $TypeSuccess
    } else {
        $bc_Repository = $pCurrentRepository
    } # else if ( $bc_browse.ShowDialog() -eq "OK" )

    return $bc_Repository
} # Function Browse_CommonRepo

#=====================================================================================
<#
    Function Import_Repository
        If this is an existing repository with filters, show the contents in the grid
            and update INI file about imported repository

        else, populate grid from INI's platform list

        Paramaters: $pGrid = pointer to model grid in UI
                    $pRepoFolder
                    $pModelsTable               --- current models table to use if nothing in repository
#>
Function Import_Repository {
    [CmdletBinding()]
	param( $pGrid, $pRepoFolder, $pModelsTable )                                

    CMTraceLog -Message "> Import_Repository($($pRepoFolder))" -Type $TypeNorm

    Empty_Grid $pGrid

    if ( [string]::IsNullOrEmpty($pRepoFolder ) ) {
        CMTraceLog -Message  "... No Repository to import" -Type $TypeWarn        
    } else {        
        # folder exists, so if it is not a HPIA repository, initialize it
        if ( Test-Path $pRepoFolder ) {            
            Try {
                Set-Location $pRepoFolder
                $ir_ProdFilters = (Get-RepositoryInfo).Filters
            } Catch {
                init_repository $pRepoFolder $true
            }
        } else {
            # Folder does not exist, so initialize it for HPIA
            init_repository $pRepoFolder $true
        } # else if ( Test-Path $pRepoFolder )
        
        Try {
            # Get the CMSL filters from the repository
            Set-Location $pRepoFolder
            CMTraceLog -Message  "... valid HPIA Repository Found" -Type $TypeNorm
            #--------------------------------------------------------------------------------
            # find out if AddOn flag files exist to add platforms to the grid
            #--------------------------------------------------------------------------------
            [array]$ir_FlagFiles = (Get-ChildItem -Path $pRepoFolder'\'$s_AddSoftware | Where-Object { ($_.Name).Length -eq 4 } ).Name  
            # ... populate grid with platform SysIDs found in the repository
            if ( $ir_FlagFiles.count -gt 0 ) { 
                foreach ( $iFlagFile in $ir_FlagFiles ) {
                    [array]$ir_ProdName = Get-HPDeviceDetails -Platform $iFlagFile
                    [void]$pGrid.Rows.Add(@( $false, $iFlagFile, $ir_ProdName[0].Name ))
                } # foreach ( $iFlagFile in $ir_FlagFiles )
            } # if ( $ir_FlagFiles.count -gt 0 )
            #--------------------------------------------------------------------------------
            if ( $ir_ProdFilters.count -eq 0 -and ($ir_FlagFiles.count -eq 0) ) {
                CMTraceLog -Message  "... no filters found, so populate from INI file (calling Populate_GridFromINI())" -Type $TypeNorm
                Populate_GridFromINI $pGrid $pModelsTable
            } else {
                #--------------------------------------------------------------------------------
                # next update the grid from in the repository filters 
                # ... first, get the (unique) list of platform SysIDs in the repository
                #--------------------------------------------------------------------------------
                [array]$ir_RepoPlatforms = (Get-RepositoryInfo).Filters.platform | Get-Unique
                # next, add each product to the grid, to then populate with the filters
                CMTraceLog -Message  "... Adding platforms to Grid from repository"
                #--------------------------------------------------------------------------------
                $ir_rowCount = $pGrid.RowCount              # are there any entries in the grid?
                for ( $i = 0 ; $i -lt $ir_RepoPlatforms.count ; $i++) {
                    $ir_SysID = $ir_RepoPlatforms[$i]
                    $ir_PlatformAdded = $false
                    for ( $ir_row = 0 ; $ir_row -lt $ir_rowCount ; $ir_row++ ) {
                        # check that the platfrom wasn't added to the grid already
                        if ( $pGrid[1,$ir_row].value -match $ir_SysID ) { 
                            $ir_PlatformAdded = $True }
                    } # for ( $ir_row = 0 ; $ir_row -lt $ir_rows ; $ir_row++ )
                    if ( -not $ir_PlatformAdded ) {
                        [array]$ir_ProdName = Get-HPDeviceDetails -Platform $ir_SysID
                        [void]$pGrid.Rows.Add(@( $false, $ir_SysID, $ir_ProdName[0].Name ))
                    } # if ( -not $ir_PlatformAdded )
                } # for ( $i = 0 ; $i -lt $ir_RepoPlatforms.count ; $i++)

                CMTraceLog -Message "... Finding filters from $($pRepoFolder) (calling Get_CommonRepofilters())" -Type $TypeNorm
                Get_CommonRepofilters $pGrid $pRepoFolder $True

                CMTraceLog -Message  "... Updating UI and INI file from filters (calling Update_UIandINISetting())" -Type $TypeNorm
                Update_UIandINISetting $pRepoFolder $True     # $True = common repo

                CMTraceLog -Message  "... Updating models in INI file (calling Update_INIModelsListFromGrid())" -Type $TypeNorm
                Update_INIModelsListFromGrid $pGrid $pRepoFolder  # $False means treat as head of individual repositories
            } # if ( $ir_ProdFilters.count -gt 0 )
            CMTraceLog -Message "< Import Repository Done" -Type $TypeSuccess
        } Catch {
            CMTraceLog -Message "< Repository Folder ($($pRepoFolder)) not initialized for HPIA" -Type $TypeWarn
        } # Catch
    } # else if ( ([string]::IsNullOrEmpty($pRepoFolder)) -or (-not (Test-Path $pRepoFolder)) )
    
} # Function Import_Repository

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
    Function Update_INIModelsListFromGrid
        Create $HPModelsTable array from current checked grid entries and updates INI file with it
        Parameters
            $pGrid
            $pRepositoryFolder
#>
Function Update_INIModelsListFromGrid {
    [CmdletBinding()]
	param( $pGrid,
        $pRepositoryFolder )            # required to find the platform ID flag files hosted in .ADDSOFTWARE

    if ( $v_DebugMode ) { CMTraceLog -Message '> Update_INIModelsListFromGrid' -Type $TypeNorm }

    $ui_INIFilePath = $Script:IniFIleFullPath
    # -------------------------------------------------------------------
    # create list of models from the grid - assumes GUI grid is populated
    # -------------------------------------------------------------------
    $ui_ModelsList = @()
    for ( $iRow = 0; $iRow -lt $pGrid.RowCount; $iRow++ ) {
        $ui_ModelId = $pGrid[1,$iRow].Value             # column 1 has the Platform SySID
        $ui_ModelName = $pGrid[2,$iRow].Value           # column 2 has the Model name
        $ui_ModelsList += "`n`t@{ ProdCode = '$($ui_ModelId)'; Model = '$($ui_ModelName)' }" # add the entry to the model list       
    } # for ( $iRow = 0; $iRow -lt $pGrid.RowCount; $iRow++ )
    if ( $v_DebugMode ) { CMTraceLog -Message "... adding:  $($ui_ModelsList)" -Type $TypeNorm }
    CMTraceLog -Message '... Created HP Models List' -Type $TypeNorm
    # ---------------------------------------------------------
    # Now, replace HPModelTable in INI file with list from grid
    # ---------------------------------------------------------
    if ( Test-Path $ui_INIFilePath ) {
        CMTraceLog -Message "... Updating Models list (INI File $ui_INIFilePath)" -Type $TypeNorm
        $ui_ListHeader = '\$HPModelsTable = .*'
        $ui_ProdsEntry = '.*@{ ProdCode = .*'
        # remove existing model entries from file
        Set-Content -Path $ui_INIFilePath -Value (get-content -Path $ui_INIFilePath | 
            Select-String -Pattern $ui_ProdsEntry -NotMatch)
        # add new model lines        
        (get-content $ui_INIFilePath) -replace $ui_ListHeader, "$&$($ui_ModelsList)" | 
            Set-Content $ui_INIFilePath
    } else {
        CMTraceLog -Message " ... INI file not updated - didn't find" -Type $TypeWarn
    } # else if ( Test-Path $lRepoLogFile )

    if ( $v_DebugMode ) { CMTraceLog -Message '< Update_INIModelsListFromGrid' -Type $TypeNorm }
    
} # Function Update_INIModelsListFromGrid 

#=====================================================================================
<#
    Function Populate_GridFromINI
    This is the MAIN function with a Gui that sets things up for the user
#>
Function Populate_GridFromINI {
    [CmdletBinding()]
	param( $pGrid,
            $pModelsTable )

    CMTraceLog -Message "... populating Grid from INI file's `$HPModels list" -Type $TypeWarn
    <#        
        @{ ProdCode = '8723'; Model = 'HP ZBook Firefly 14 G7 Mobile Workstation' }
        @{ ProdCode = '8ABB 896D'; Model = 'HP EliteBook 840 14 inch G9 Notebook PC' }
    #> 
    $pModelsTable | 
        ForEach-Object {
            # handle case of multiple ProdCodes in model entry
            foreach ( $iSysID in $_.ProdCode.split(' ') ) {
                [void]$pGrid.Rows.Add( @( $False, $iSysID, $_.Model) )   # populate checkmark, ProdId, Model Name
                CMTraceLog -Message "... adding $($iSysID):$($_.Model)" -Type $TypeNorm
            } # foreach ( $iSysID in $_.ProdCode.split(' ') )
        } # ForEach-Object

    $pGrid.Refresh()

} # Populate_GridFromINI

#=====================================================================================
<#
    Function Update_UIandINISetting
    Update UI elements (selected path, default selection, and INI path settings)
#>
Function Update_UIandINISetting {
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

} # Function Update_UIandINISetting

#=====================================================================================
<#
    Function Create_AddOnFlagFile
        Creates and updates the Pltaform ID flag file in the repository's .ADDSOFTWARE folder        
    Parameters                 
        $pPath
        $pSysID
        $pModelName
        [array]$pAddOns        
        $pCommonFlag                        # is this a common repo?
#>
Function Create_AddOnFlagFile {
    [CmdletBinding()]
	param( $pPath,
            $pSysID,                        # 4-digit hex-code platform/motherboard ID
            $pModelName,
            [array]$pAddOns,                # $Null means empty flag file (nothing to add)
            $pCommonFlag)

    # build the flag file name path, and the file backup (when AddOn col is unchecked)
    switch ( $pCommonFlag ) {
        $true { $ca_FlagFilePath = $pPath+'\'+$s_AddSoftware+'\'+$pSysID } # $true
        $false { $ca_FlagFilePath = $pPath+'\'+$pModelName+'\'+$s_AddSoftware+'\'+$pSysID } # $false
    } # switch ( $pCommonFlag )

    # create the file and add selected Softpaqs content
    if ( -not (Test-Path $ca_FlagFilePath) ) { New-Item $ca_FlagFilePath -Force > $null } 
    if ( $pAddOns ) {
        For ($i=0; $i -lt $pAddOns.count; $i++) { $pAddOns[$i] | Out-File -FilePath $ca_FlagFilePath -Append }
    } # if ( $null -ne $pAddOns )

    return $pAddOns.count
} # Function Create_AddOnFlagFile

#=====================================================================================
<#
    Function Manage_AddOnFlagFile
        Creates and updates the Pltaform ID flag file in the repository's .ADDSOFTWARE folder
        If Flag file exists and user 'unselects' AddOns, we move the file to a backup file 'ABCD_bk'
    Parameters                 
        $pPath
        $pSysID
        $pModelName
        [array]$pAddOns
        $pUseFile                           # flag:$True to create the flag file, $False to remove (renamed)
        $pCommonFlag                        # is this a common repo?
#>
Function Manage_AddOnFlagFile {
    [CmdletBinding()]
	param( $pPath,
            $pSysID,                        # 4-digit hex-code platform/motherboard ID
            $pModelName,
            [array]$pAddOns,                # $Null means empty flag file (nothing to add)
            $pUseFile,                      # $True = create file, $False = move to backup and remove
            $pCommonFlag) 

    if ( $v_DebugMode ) { CMTraceLog -Message "... > Manage_AddOnFlagFile $pSysID : $pModelName" -Type $TypeNorm  }
    
    # ---------------------------------------------------------------------------------------
    # build the flag file name path, and the file backup (when AddOn col is unchecked)
    # ---------------------------------------------------------------------------------------
    switch ( $pCommonFlag ) {
        $true { $ma_FlagFilePath = $pPath+'\'+$s_AddSoftware+'\'+$pSysID }
        $false { $ma_FlagFilePath = $pPath+'\'+$pModelName+'\'+$s_AddSoftware+'\'+$pSysID }
    } # switch ( $pCommonFlag )
    $ma_FlagFilePathBK = $ma_FlagFilePath+'_BK'
    #$ma_RepoPathSysID = $pPath+'\'+$pSysID+'_'+$pModelName+'\'+$s_AddSoftware+'\'+$pSysID
    #$ma_RepoPathSysIDBK = $ma_RepoPathSysID+'_BK'

    if ( $pUseFile ) {
        if ( -not (Test-Path $ma_FlagFilePath) ) {
            if ( Test-Path $ma_FlagFilePathBK ) {
                Move-Item $ma_FlagFilePathBK $ma_FlagFilePath -Force
            } else {
                CMTraceLog -Message "... calling Create_AddOnFlagFile()" -Type $TypeNorm
                Create_AddOnFlagFile $pPath $pSysID $pModelName $pAddOns $pCommonFlag
            } # else if ( Test-Path $ma_FlagFilePathBK )
        } # if ( -not (Test-Path $ma_FlagFilePath) )
        CMTraceLog -Message "... calling Edit_AddonFlagFile() - for additional Softpaqs" -Type $TypeNorm
        Edit_AddonFlagFile $ma_FlagFilePath      
    } else {
        if ( Test-Path $ma_FlagFilePath ) { 
            Move-Item $ma_FlagFilePath $ma_FlagFilePathBK -Force 
        }
    } # else if ( $pUseFile )
 
    # find out how many entries in the flag file
    [int]$ma_NumOfEntries = 0
    if ( Test-Path $ma_FlagFilePath ) {
        $ma_NumOfEntries = (Get-Content $ma_FlagFilePath | Measure-Object).Count
    } elseif ( Test-Path $ma_FlagFilePathBK ) {
        $ma_NumOfEntries = (Get-Content $ma_FlagFilePathBK | Measure-Object).Count
    } # else if ( Test-Path $ma_FlagFilePath )

    if ( $v_DebugMode ) { CMTraceLog -Message '... < Manage_AddOnFlagFile' -Type $TypeNorm }
    CMTraceLog -Message $lMsg -Type $TypeNorm
    return $ma_NumOfEntries

} # Manage_AddOnFlagFile

#=====================================================================================
<#
    Function Edit_AddonFlagFile

    manage the Softpaq Addons in the flag file
    called to confirm what entries to be maintained in the ID flag file
#>
Function Edit_AddonFlagFile { 
    [CmdletBinding()]
	param( $pFlagFilePath )

    #$pSysID = Split-Path $pFlagFilePath -leaf
    $ea_FileContents = Get-Content $pFlagFilePath

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
            if ( $SoftpaqEntry.Text -in $ea_FileContents ) {
                CMTraceLog -Message '... item already included' -Type $TypeNorm
            } else {
                $EntryList.items.add($SoftpaqEntry.Text)
                $ea_entries = @()
                foreach ( $iEntry in $EntryList.items ) { $ea_entries += $iEntry }
                Set-Content $pFlagFilePath -Value $ea_entries      # reset the file with needed AddOns
            } # else if ( $SoftpaqEntry.Text -in $ea_FileContents )
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

    foreach ( $iSoftpaqName in $ea_FileContents ) { $EntryList.items.add($iSoftpaqName) }

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
            $ea_entries = @()
            foreach ( $iEntry in $EntryList.items ) { $ea_entries += $iEntry }
            Set-Content $pFlagFilePath -Value $ea_entries      # reset the file with needed AddOns
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
    $lRet = $SoftpaqsForm.ShowDialog()

} # Function Edit_AddonFlagFile

#=====================================================================================
<#
    Function Remove_Platform
    Removes entries from the file system associated with this platform
    The grid entry is removed by the caller of this function
    Repository is renamed adding '_BK'
    If multiple flag AddOn files, the removed platform flag file is removed leaving the repository intact
    if the model is part of a Common repository, all that's done is remove the flag file
        ... next Sync will cleanup Softpaqs from the repo
    parameters
        $pGrid
        pGridIndex
        $pRepositoryFolder
        $pCommonFlag
#>
Function Remove_Platform {
    [CmdletBinding()]
        param( $pGrid, 
                $pGridIndex,
                $pRepositoryFolder, 
                $pCommonFlag )

    $rp_SysID = $pGrid.Rows[$pGridIndex].Cells[1].Value     # obtain the platform SysID from the grid row (2nd col)

    switch ( $pCommonFlag ) {
        $true {
            # Since the repo may contain multiple platforms, we just remove the flag file here
            $rp_FlagFilePath = $pRepositoryFolder+'\'+$s_AddSoftware+'\'+$rp_SysID  
            'common - searching for flag file: '+$rp_FlagFilePath | Out-Host
            if ( Test-Path $rp_FlagFilePath ) {
                '-- found: '+$rp_FlagFilePath | Out-Host
                Remove-Item -Path $rp_FlagFilePath -confirm
            } # if ( Test-Path $rp_FlagFilePath )
        } # $true
        $false {
            # here we are dealing with individual repository,  possibly a repo with multiple SysIDs
            # ... e.g. there could be >1 SysID AddOns flag files to deal with, in that case, we just
            # ...     remove the flag file, keeping the repository intact for other SysIDs
            # ... if only 1 flag file exists, then we will rename the repo folder adding '_BK' to the name
            $rp_SysName = $pGrid.Rows[$pGridIndex].Cells[2].Value
            $rp_RepoPath = $pRepositoryFolder+'\'+$rp_SysName
            
            $rp_FlagFilePath = $rp_RepoPath+'\'+$s_AddSoftware+'\'+$rp_SysID    # this is the model we are removing from the grid

            if ( Test-Path $rp_RepoPath ) {
                # flag file name has 4 characters - find all of them, in case there are multiple SysIDs in the grid
                [array]$rp_FlagFiles = Get-ChildItem $rp_RepoPath'\'$s_AddSoftware | 
                    Where-Object {$_.BaseName.Length -eq 4} 

                if ( $rp_FlagFiles.count -gt 1  ) {
                    Remname-Item -Path $rp_RepoPath -NewName $rp_SysName'_BK' -confirm          # just rename the repository file
                } else {
                    if ( $rp_FlagFiles.count -eq 1 -and ($rp_FlagFiles[0] -like $rp_SysID) ) {
                        Remove-Item -Path $rp_FlagFilePath -confirm                             # remove the flag file only
                    }
                }               
            } # if ( Test-Path $rp_RepoPath )

        } # $false
    } # switch ( $pCommonFlag )

    $pGrid.Rows.RemoveAt($pGridIndex)

} # Function Remove_Platform

#=====================================================================================
<#
    Function Remove_Platforms
    Remove the model from the list and the respotory info, after confirmation
    parameters
        $pGrid, 
        $pRepositoryFolder, 
        $pCommonFlag
#>
Function Remove_Platforms {
    [CmdletBinding()]
	param( $pGrid, $pRepositoryFolder, $pCommonFlag )

    $rp_removeList = @()
    # confirm there are selected entries/models to remove from the list
    for ( $irow = 0 ; $irow -lt $pGrid.RowCount ; $irow++ ) {
        if ( $pGrid.Rows[$irow].Cells[0].Value ) {
            $rp_removeList += @{
                    SysID = $pGrid.Rows[$irow].Cells[1].Value
                    SysName = $pGrid.Rows[$irow].Cells[2].Value 
                }     
            $rp_removeList +=$rp_platform
        } # if ( $pGrid.Rows[$irow].Cells[0].Value )
    } # for ( $row = 0 ; $row -lt $pGrid.RowCount ; $row++ )

    if ( $rp_removeList ) {
        $rp_Msg = "Remove the selected entries from the list and delete associated repository information"
    } else {
        $rp_Msg = "There are no selected entries to remove"
    }    
    #$lAsk = [System.Windows.MessageBox]::Show($rp_Msg,"Remove Model",1)
    $lAsk = AskUser $rp_Msg 1 # 1="OKCancel"
    if ( $lAsk -eq 'Ok') {
        'REMOVING entry and repository information' | Out-Host
        for ( $irow = $pGrid.RowCount-1 ; $irow -ge 0 ; $irow-- ) {
            if ( $pGrid.Rows[$irow].Cells[0].Value ) {
                Remove_Platform $pGrid $irow $pRepositoryFolder $pCommonFlag
                #$pGrid.Rows.RemoveAt($irow)
 
            } # if ( $pGrid.Rows[$irow].Cells[0].Value )
        } # for ( $irow = 0 ; $irow -lt $pGrid.RowCount ; $irow++ )
        
        CMTraceLog -Message '... calling Update_INIModelsListFromGrid()' -Type $TypeNorm 
        Update_INIModelsListFromGrid $pGrid $pRepositoryFolder            
    } else {
        $rp_removeList = $null
    } # if ( $lAsk -eq 'Ok')
 
    return $rp_removeList
} # Function Remove_Platforms

#=====================================================================================
<#
    Function Add_Platform
    Ask the user to select a new device to add to the list
#>
Function Add_Platform {
    [CmdletBinding()]
	param( $pGrid, $pSoftpaqList, $pRepositoryFolder, $pCommonFlag )

    Set-Variable ap_FormWidth -Option Constant -Value 400
    Set-Variable ap_FormHeigth -Option Constant -Value 400
    Set-Variable ap_Offset -Option Constant -Value 20
    Set-Variable ap_FieldHeight -Option Constant -Value 20
    Set-Variable ap_PathFieldLength -Option Constant -Value 200

    $EntryForm = New-Object System.Windows.Forms.Form
    $EntryForm.MaximizeBox = $False ; $EntryForm.MinimizeBox = $False #; $EntryForm.ControlBox = $False
    $EntryForm.Text = "Find a Device to Add"
    $EntryForm.Width = $ap_FormWidth ; $EntryForm.height = 400 ; $EntryForm.Autosize = $true
    $EntryForm.StartPosition = 'CenterScreen'
    $EntryForm.Topmost = $true

    # ------------------------------------------------------------------------
    # find and add model entry
    # ------------------------------------------------------------------------
    $EntryId = New-Object System.Windows.Forms.Label
    $EntryId.Text = "Name"
    $EntryId.location = New-Object System.Drawing.Point($ap_Offset,$ap_Offset) # (from left, from top)
    $EntryId.Size = New-Object System.Drawing.Size(60,20)                   # (width, height)

    $EntryModel = New-Object System.Windows.Forms.TextBox
    $EntryModel.Text = ""       # start w/INI setting
    $EntryModel.Multiline = $false 
    $EntryModel.location = New-Object System.Drawing.Point(($ap_Offset+70),($ap_Offset-4)) # (from left, from top)
    $EntryModel.Size = New-Object System.Drawing.Size($ap_PathFieldLength,$ap_FieldHeight)# (width, height)
    $EntryModel.ReadOnly = $False
    $EntryModel.Name = "Model Name"
    $EntryModel.add_MouseHover($ShowHelp)
    $SearchButton = New-Object System.Windows.Forms.Button
    $SearchButton.Location = New-Object System.Drawing.Point(($ap_PathFieldLength+$ap_Offset+80),($ap_Offset-6))
    $SearchButton.Size = New-Object System.Drawing.Size(75,23)
    $SearchButton.Text = 'Search'
    $SearchButton_AddClick = {
        if ( $EntryModel.Text ) {
            $AddEntryList.Items.Clear()
            $ap_Models = Get-HPDeviceDetails -Like -Name $EntryModel.Text    # find all models matching entered text
            foreach ( $iModel in $ap_Models ) { 
                [void]$AddEntryList.Items.Add($iModel.SystemID+'_'+$iModel.Name) 
            }
        } # if ( $EntryModel.Text )
    } # $SearchButton_AddClick =
    $SearchButton.Add_Click( $SearchButton_AddClick )

    $lListHeight = $ap_FormHeigth/2-60
    $AddEntryList = New-Object System.Windows.Forms.ListBox
    $AddEntryList.Name = 'Entries'
    $AddEntryList.Autosize = $false
    $AddEntryList.location = New-Object System.Drawing.Point($ap_Offset,60)  # (from left, from top)
    $AddEntryList.Size = New-Object System.Drawing.Size(($ap_FormWidth-60),$lListHeight) # (width, height)
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
    $SpqLabel.location = New-Object System.Drawing.Point($ap_Offset,($lListHeight+60)) # (from left, from top)
    $SpqLabel.Size = New-Object System.Drawing.Size(70,60)                   # (width, height)

    $AddSoftpaqList = New-Object System.Windows.Forms.ListBox
    $AddSoftpaqList.Name = 'Softpaqs'
    $AddSoftpaqList.Autosize = $false
    $AddSoftpaqList.SelectionMode = 'MultiExtended'
    $AddSoftpaqList.location = New-Object System.Drawing.Point(($ap_Offset+70),($lListHeight+60))  # (from left, from top)
    $AddSoftpaqList.Size = New-Object System.Drawing.Size(($ap_FormWidth-130),($lListHeight-40)) # (width, height)

    # ------------------------------------------------------------------------
    # show the dialog, and once user preses OK, add the model and create the flag file for addons
    # ------------------------------------------------------------------------
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(($ap_FormWidth-120),($ap_FormHeigth-80))
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(($ap_FormWidth-200),($ap_FormHeigth-80))
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::CANCEL

    $EntryForm.AcceptButton = $okButton
    $EntryForm.CancelButton = $cancelButton
    $EntryForm.Controls.AddRange(@($EntryId,$EntryModel,$SearchButton,$AddEntryList,$SpqLabel,$AddSoftpaqList, $cancelButton, $okButton))

    $ap_Result = $EntryForm.ShowDialog()

    if ($ap_Result -eq [System.Windows.Forms.DialogResult]::OK) {        
        
        $ap_SelectedSysID = $AddEntryList.SelectedItem.substring(0,4)
        $ap_SelectedModel = $AddEntryList.SelectedItem.substring(5)  # name is after 'SysID_'
        
        [array]$ap_SelectedEntry = Get-HPDeviceDetails -Like -Name $AddEntryList.SelectedItem

        # see if model is already in the grid, and avoid it, otherwise, add it
        for ( $iRow = 0; $iRow -lt $pGrid.RowCount; $iRow++ ) {
            $ap_CurrEntrySysID = $pGrid.Rows[$iRow].Cells[1].value
            if ( $ap_CurrEntrySysID -like $ap_SelectedSysID ) {
                #$lRes = [System.Windows.MessageBox]::Show("This model is already in the Grid","Add Model to Grid",0)    # 1 = "OKCancel" ; 4 = "YesNo"
                $lRes = AskUser "This model is already in the Grid" 0 # 0="OK"
                return $null
            }
        } # for ( $iRow = 0; $iRow -lt $pGrid.RowCount; $iRow++ )

        # add model to UI grid and initialize it as an HPIA repository
        [void]$pGrid.Rows.Add( @( $False, $ap_SelectedSysID, $ap_SelectedModel) )

        # get a list of selected additional softpaqs to add
        CMTraceLog -Message '... calling Create_AddOnFlagFile()' -Type $TypeNorm 
        $ap_numEntries = Create_AddOnFlagFile $pRepositoryFolder $ap_SelectedSysID $ap_SelectedModel $AddSoftpaqList.SelectedItems $pCommonFlag

        # check the AddOns cell if entries in flag file
        if ( $ap_numEntries -gt 0 ) {
            $pGrid.Rows[($pGrid.RowCount-1)].Cells['AddOns'].Value = $True
        }
        CMTraceLog -Message '... calling Update_INIModelsListFromGrid()' -Type $TypeNorm 
        Update_INIModelsListFromGrid $pGrid $pRepositoryFolder
        return $ap_SelectedEntry
   
    } # if ($ap_Result -eq [System.Windows.Forms.DialogResult]::OK)

} # Function Add_Platform

########################################################################################

Function AskUser {
    [CmdletBinding()]
        param( $pText, $pAskType )
    # 0='OK', 1="OKCancel", 4="YesNo"
    switch ( $pAskType ) {
        0 { $au_askbutton2 = 'OK' }
        1 { $au_askbutton1 = 'OK' ; $au_askbutton2 = 'Cancel' }
        4 { $au_askbutton1 = 'Yes' ; $au_askbutton2 = 'No' }
    }
    
    $au_FormWidth = 300 ; $au_FormHeight = 200

    $au_form = New-Object System.Windows.Forms.Form
    $au_form.Text = ""
    $au_form.Width = $au_FormWidth ; $au_form.height = $au_FormHeight #; $au_form.Autosize = $true
    $au_form.StartPosition = 'CenterScreen'
    $au_form.Topmost = $true

    $au_MsgBox = New-Object System.Windows.Forms.TextBox
    $au_MsgBox.BorderStyle = 0
    $au_MsgBox.Text = $pText       # start w/INI setting
    $au_MsgBox.Multiline = $true 
    $au_MsgBox.location = New-Object System.Drawing.Point(20,20) # (from left, from top)
    $au_MsgBox.Size = New-Object System.Drawing.Size(($au_FormWidth-60),($au_FormHeight-120))# (width, height)
    #$au_MsgBox.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",($TextBox.Font.Size+4),[System.Drawing.FontStyle]::Regular)
    $au_MsgBox.Font = New-Object System.Drawing.Font("Verdana",($TextBox.Font.Size+4),[System.Drawing.FontStyle]::Regular)
    $au_MsgBox.TabStop = $false
    $au_MsgBox.ReadOnly = $true
    $au_MsgBox.Name = "Ask User"

    #----------------------------------------------------------------------------------
    # Create Buttons at the bottom of the dialog
    #----------------------------------------------------------------------------------
    $au_button1 = New-Object System.Windows.Forms.Button
    $au_button1.Text = $au_askbutton1
    $au_button1.Location = New-Object System.Drawing.Point(($au_FormWidth-200),($au_FormHeight-80))    # (from left, from top)
    $au_button1.add_click( { $au_return = $au_askbutton1 ; $au_form.Close() ; $au_form.Dispose() } )

    $au_button2 = New-Object System.Windows.Forms.Button
    $au_button2.Text = $au_askbutton2
    $au_button2.Location = New-Object System.Drawing.Point(($au_FormWidth-120),($au_FormHeight-80))    # (from left, from top)
    $au_button2.add_click( { $au_return = $au_askbutton2 ; $au_form.Close() ; $au_form.Dispose() } )
    if ( $pAskType -eq 0 ) {
        $au_form.Controls.AddRange(@($au_MsgBox,$au_button2))
    } else {
        $au_form.Controls.AddRange(@($au_MsgBox,$au_button1,$au_button2))
    }

    $au_form.ShowDialog() | Out-Null
    return $au_return 
} # Function AskUser

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
    # Create Add accesories category when selecting 'Driver'
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating Add accesories category Checkbox' }
    $AddAccessoriesCheckbox = New-Object System.Windows.Forms.CheckBox
    $AddAccessoriesCheckbox.Text = 'Add Accesories with Driver'
    $AddAccessoriesCheckbox.Name = 'Add Accesories'
    #$AddAccessoriesCheckbox.add_MouseHover($ShowHelp)
    $AddAccessoriesCheckbox.Autosize = $true
    $AddAccessoriesCheckbox.Location = New-Object System.Drawing.Point(($LeftOffset+160),($TopOffset+42))   # (from left, from top)
    $AddAccessoriesCheckbox.add_Click({
        if ( $AddAccessoriesCheckbox.checked ) {
            $Script:s_AddAccesories = $true
        }    
    })
    $CM_form.Controls.AddRange(@($AddAccessoriesCheckbox))

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
    #$OSMessage = 'Check and confirm the OS/OS Version selected for this Repository is correct'
    $OSMessage = 'Confirm the selected OS and OS Version to sync this Repository is correct'
    $OSHeader = 'Imported Repository'

    if ( $v_DebugMode ) { Write-Host 'creating Shared radio button' }
    $IndividualRadioButton = New-Object System.Windows.Forms.RadioButton
    $IndividualRadioButton.Location = '10,14'
    $IndividualRadioButton.Add_Click( {
            $Script:v_CommonRepo = $False
            $find = "^[\$]v_CommonRepo"
            $replace = "`$v_CommonRepo = `$$Script:v_CommonRepo" 
            Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - INI setting set: ''$($replace)''"
            Import_RootedRepos $dataGridView $IndividualPathTextField.Text            
            #[System.Windows.MessageBox]::Show($OSMessage,$OSHeader,0)    # 0='OK', 1="OKCancel", 4="YesNo"
            AskUser $OSMessage 0 # 0="OK"            
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
    $individualBrowse = New-Object System.Windows.Forms.Button
    $individualBrowse.Width = 60
    $individualBrowse.Text = 'Browse'
    $individualBrowse.Location = New-Object System.Drawing.Point(($LeftOffset+$PathFieldLength+$labelWidth+40),($TopOffset-5))
    $individualBrowse_Click = {
        $Script:v_CommonRepo = $false
        $find = "^[\$]v_CommonRepo"
        $replace = "`$v_CommonRepo = `$$Script:v_CommonRepo" 
        Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"
        $mf_Repository = Browse_IndividualRepos $dataGridView $IndividualPathTextField.Text
        if ( $mf_repository ) {
            $IndividualRadioButton.checked = $true
            $IndividualPathTextField.Text = $mf_repository            
            $IndividualPathTextField.BackColor = $BackgroundColor
            $CommonPathTextField.BackColor = ''
         }
    } # $individualBrowse_Click
    $individualBrowse.add_Click($individualBrowse_Click)

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
        Import_Repository $dataGridView $CommonPathTextField.Text $Script:HPModelsTable
        #[System.Windows.MessageBox]::Show($OSMessage,$OSHeader,0)    # 1 = "OKCancel" ; 4 = "YesNo"
        AskUser $OSMessage 0
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
        $Script:v_CommonRepo = $True
        $find = "^[\$]v_CommonRepo"
        $replace = "`$v_CommonRepo = `$$Script:v_CommonRepo" 
        Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"
        $mf_repository = Browse_CommonRepo $dataGridView $CommonPathTextField.Text $Script:HPModelsTable
        if ( $mf_repository ) {
            $CommonPathTextField.Text = $mf_repository
            $CommonPathTextField.BackColor = $BackgroundColor
            $IndividualPathTextField.BackColor = ''
         }
    } # $commonBrowse_Click
    $commonBrowse.add_Click($commonBrowse_Click)

    $PathsGroupBox.Controls.AddRange(@($IndividualPathTextField, $SharePathLabel,$IndividualRadioButton, $individualBrowse))
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
                        Manage_AddOnFlagFile $lRepoPath $datagridview[1,$row].value $datagridview[2,$row].value $Null $CellNewState $Script:v_CommonRepo
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
    $CMModelsGroupBox.location = New-Object System.Drawing.Point($LeftOffset,($TopOffset+70))     # (from left, from top)
    $CMModelsGroupBox.Size = New-Object System.Drawing.Size(($ListViewWidth+20),($ListViewHeight+30))       # (width, height)
    $CMModelsGroupBox.text = "HP Models / Repository Category Filters"

    $CMModelsGroupBox.Controls.AddRange(@($dataGridView))

    $CM_form.Controls.AddRange(@($CMModelsGroupBox))

    #----------------------------------------------------------------------------------
    # Add a Use List button
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating a Refresh List button' }
    $RefreshGridButton = New-Object System.Windows.Forms.Button
    $RefreshGridButton.Location = New-Object System.Drawing.Point($LeftOffset,($FormHeight-420))    # (from left, from top)
    $RefreshGridButton.Text = 'Use INI List'
    $RefreshGridButton.Name = 'Use INI List'
    $RefreshGridButton.AutoSize=$true
    $RefreshGridButton.add_MouseHover($ShowHelp)

    $RefreshGridButton_Click={
        #CMTraceLog -Message 'Pupulating HP Models List from INI file $HPModels' -Type $TypeNorm
        Empty_Grid $dataGridView
        Populate_GridFromINI $dataGridView $Script:HPModelsTable
    } # $RefreshGridButton_Click={

    $RefreshGridButton.add_Click($RefreshGridButton_Click)

    $CM_form.Controls.Add($RefreshGridButton)
  
    #----------------------------------------------------------------------------------
    # Add a list filters button
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating a list filters button' }
    $ListFiltersdButton = New-Object System.Windows.Forms.Button
    $ListFiltersdButton.Location = New-Object System.Drawing.Point(($LeftOffset+93),($FormHeight-427))    # (from left, from top)
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
    $AddModelButton.Width = 80 ; $AddModelButton.Height = 35
    $AddModelButton.Location = New-Object System.Drawing.Point(($LeftOffset+190),($FormHeight-427))    # (from left, from top)
    $AddModelButton.Text = 'Add Model'
    $AddModelButton.Name = 'Add Model'

    $AddModelButton.add_Click( { 
        if ( $Script:v_CommonRepo ) { $mf_RepoPath = $CommonPathTextField.Text } else { $mf_RepoPath = $IndividualPathTextField.Text }
        $mf_ModelArray = Add_Platform $DataGridView $Script:v_Softpaqs $mf_RepoPath $Script:v_CommonRepo
        if ( $null -ne $mf_ModelArray ) {
            $Script:s_ModelAdded = $True
            CMTraceLog -Message "$($mf_ModelArray.ProdCode):$($mf_ModelArray.Name) added to list" -Type $TypeNorm
        } 
    } ) # $AddModelButton.add_Click(
    $CM_form.Controls.Add($AddModelButton)
    #----------------------------------------------------------------------------------
    # Add a Remove Model button
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating a Add Model button' }
    $RemoveModelButton = New-Object System.Windows.Forms.Button
    $RemoveModelButton.Width = 80 ; $RemoveModelButton.Height = 35
    $RemoveModelButton.Location = New-Object System.Drawing.Point(($LeftOffset+270),($FormHeight-427))    # (from left, from top)
    $RemoveModelButton.Text = 'Remove Model'
    $RemoveModelButton.Name = 'Remove Model'
    $RemoveModelButton.add_Click( { 
        if ( $Script:v_CommonRepo ) { $mf_RepoPath = $CommonPathTextField.Text } else { $mf_RepoPath = $IndividualPathTextField.Text }
        (Remove_Platforms $DataGridView $mf_RepoPath $Script:v_CommonRepo) | foreach {
            $Script:s_ModelRemoved = $True
            if ( $_ ) { CMTraceLog -Message "Platform removed from list: $($_.SysID) / $($_.SysName) " -Type $TypeNorm }
        } # $mf_ModelArray | foreach
    } ) # $AddModelButton.add_Click(
    $CM_form.Controls.Add($RemoveModelButton)

    #----------------------------------------------------------------------------------
    # Create New Log checkbox
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating New Log checkbox' }
    $NewLogCheckbox = New-Object System.Windows.Forms.CheckBox
    $NewLogCheckbox.Text = 'Start New Log'
    $NewLogCheckbox.Autosize = $true
    $NewLogCheckbox.Location = New-Object System.Drawing.Point(($LeftOffset+370),($FormHeight-430))
    $Script:NewLog = $NewLogCheckbox.Checked
    $NewLogCheckbox_Click = {
        if ( $NewLogCheckbox.checked ) { $Script:NewLog = $true } else { $Script:NewLog = $false }
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
    $LogPathField.location = New-Object System.Drawing.Point(($LeftOffset+368),($FormHeight-408)) # (from left, from top)
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
    $buttonSync.Location = New-Object System.Drawing.Point(($FormWidth-130),($FormHeight-427))
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
                Sync_CommonRepository $dataGridView $CommonPathTextField.Text $lCheckedListArray #$Script:s_ModelAdded
                if ( $Script:s_ModelAdded ) { Update_INIModelsListFromGrid $dataGridView $CommonPathTextField }   # $True = head of common repository
            } else {
                sync_individualRepositories $dataGridView $IndividualPathTextField.Text $lCheckedListArray $Script:s_ModelAdded
                if ( $Script:s_ModelAdded ) { Update_INIModelsListFromGrid $dataGridView $IndividualPathTextField.Text }   # $False = head of individual repos
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
    $CMGroupAll.location = New-Object System.Drawing.Point($LeftOffset,($FormHeight-385))         # (from left, from top)
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
        #$lRes = [System.Windows.MessageBox]::Show("Create or Update HPIA Package in CM?","HP Image Assistant",1)    # 1 = "OKCancel" ; 4 = "YesNo"
        $lRes = AskUser "Create or Update HPIA Package in CM?" 1 # 1 = 'OKCancel'
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
    $TextBox.location = New-Object System.Drawing.Point($LeftOffset,($FormHeight-320))  # (from left, from top)
    $TextBox.Size = New-Object System.Drawing.Size(($FormWidth-60),($FormHeight/2-140)) # (width, height)

    $TextBoxFontDefault =  $TextBox.Font    # save the default font
    $TextBoxFontDefaultSize = $TextBox.Font.Size

    $CM_form.Controls.AddRange(@($TextBox))
 
    #----------------------------------------------------------------------------------
    # Add a clear TextBox button
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating a clear textbox checkmark' }
    $ClearTextBox = New-Object System.Windows.Forms.Button
    $ClearTextBox.Location = New-Object System.Drawing.Point($LeftOffset,($FormHeight-52))    # (from left, from top)
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
    $DebugCheckBox.location = New-Object System.Drawing.Point(($LeftOffset+120),($FormHeight-50))   # (from left, from top)
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
    $TextBoxFontdButtonDec.Location = New-Object System.Drawing.Point(($LeftOffset+240),($FormHeight-52))    # (from left, from top)
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
    $TextBoxFontdButtonInc.Location = New-Object System.Drawing.Point(($LeftOffset+310),($FormHeight-52))    # (from left, from top)
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
    if ( $v_DebugMode ) { Write-Host 'creating a TextBox TEST Connection button' }
    $TextBoxNetTestButton = New-Object System.Windows.Forms.Button
    $TextBoxNetTestButton.Location = New-Object System.Drawing.Point(($FormWidth-300),($FormHeight-52))    # (from left, from top)
    $TextBoxNetTestButton.Text = 'Test HP Connection'
    $TextBoxNetTestButton.AutoSize = $true

    $TextBoxNetTestButton_Click={
        
        CMTraceLog -Message '> Testing TCP connections to HP download servers... Please wait' -Type $TypeNorm

        CMTraceLog -Message '... { hpia.hpcloud.hp.com }' -Type $TypeNoNewline
        $r = Test-NetConnection hpia.hpcloud.hp.com -CommonTCPPort HTTP 2>&1 3>&1 #| Tee-Object -Variable out_content
        if ( $r.Count -eq 1 ) {
            CMTraceLog -Message " HTTP:/$($r.RemoteAddress) : $($r.TcpTestSucceeded)" -Type $TypeWarn
        } else {
            CMTraceLog -Message " HTTP:/$($r[1].RemoteAddress):$($r[1].TcpTestSucceeded)" -Type $TypeWarn
            CMTraceLog -Message "... $($r[0]) -- " -Type $TypeWarn
        }       
        $r = Test-Connection hpia.hpcloud.hp.com -Count 1 -Quiet
        CMTraceLog -Message "...    ICMP Test: $r" -Type $TypeWarn

        CMTraceLog -Message '... { ftp.hp.com }' -Type $TypeNoNewline
        $r = Test-NetConnection ftp.hp.com -CommonTCPPort HTTP 2>&1 3>&1 #| Tee-Object -Variable out_content
        if ( $r.Count -eq 1 ) {
            CMTraceLog -Message " HTTP:/$($r.RemoteAddress) : $($r.TcpTestSucceeded)" -Type $TypeWarn
        } else {
            $RemoteServerCount = ($r[1].RemoteAddress).count
            if ( $RemoteServerCount -eq 1 ) {
                CMTraceLog -Message " HTTP:/$($r[1].RemoteAddress) -- $($r[1].TcpTestSucceeded)" -Type $TypeWarn
                CMTraceLog -Message "... $($[0]) -- " -Type $TypeWarn                
            } else {
                for ($i=0;$i -lt $RemoteServerCount) {
                    CMTraceLog -Message " HTTP:/$($r[$i].RemoteAddress) -- $($r[$i].TcpTestSucceeded)" -Type $TypeWarn
                }
            }
        }          
        $r = Test-Connection ftp.hp.com -Count 1 -Quiet
        CMTraceLog -Message "...    ICMP Test: $r" -Type $TypeWarn

        CMTraceLog -Message '... { ftp.ext.hp.com }' -Type $TypeNoNewline
        $r = Test-NetConnection ftp.ext.hp.com -CommonTCPPort HTTP 2>&1 3>&1 #| Tee-Object -Variable out_content
         if ( $r.Count -eq 1 ) {
            CMTraceLog -Message " HTTP:/$($r.RemoteAddress) : $($r.TcpTestSucceeded)" -Type $TypeWarn
        } else {
            CMTraceLog -Message " HTTP:/$($r[1].RemoteAddress) -- $($r[1].TcpTestSucceeded)" -Type $TypeWarn
            CMTraceLog -Message "... $($r[0]) -- " -Type $TypeWarn
        }  
        $r = Test-Connection ftp.ext.hp.com -Count 1 -Quiet
        CMTraceLog -Message "...    ICMP Test: $r" -Type $TypeWarn
        #$r = Test-NetConnection ftp.ext.hp.com 2>&1 3>&1 #| Tee-Object -Variable out_content
        #CMTraceLog -Message "ICMP:/$($r.RemoteAddress) -- $($r.TcpTestSucceeded)" -Type $TypeWarn

        CMTraceLog -Message '< Network test completed' -Type $TypeSuccess
        $textbox.refresh() 
    } # $TextBoxNetTestButton_Click={

    $TextBoxNetTestButton.add_Click($TextBoxNetTestButton_Click)

   $CM_form.Controls.Add($TextBoxNetTestButton)
  
    #----------------------------------------------------------------------------------
    # Create Done/Exit Button at the bottom of the dialog
    #----------------------------------------------------------------------------------
    if ( $v_DebugMode ) { Write-Host 'creating Done/Exit Button' }
    $buttonDone = New-Object System.Windows.Forms.Button
    $buttonDone.Text = 'Exit'
    $buttonDone.Location = New-Object System.Drawing.Point(($FormWidth-140),($FormHeight-50))    # (from left, from top)

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
                    Write-Host "Common Repository Path from INI file not found: $($lCommonPath) - Will Create" -ForegroundColor Brown
                    Init_Repository $lCommonPath $True                 # $True = make it a HPIA repo
                }   
                Import_Repository $dataGridView $lCommonPath $Script:HPModelsTable
                $CommonRadioButton.Checked = $true                     # set the visual default from the INI setting
                $CommonPathTextField.BackColor = $BackgroundColor
            } else { 
                if ( [string]::IsNullOrEmpty($IndividualPathTextField.Text) ) {
                    Write-Host "Individual Repository field is empty" -ForegroundColor Red
                } else {
                    Import_RootedRepos $dataGridView $IndividualPathTextField.Text
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