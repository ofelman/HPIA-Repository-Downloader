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

$ScriptVersion = "2.05.01 (Mar 31, 2025)"

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
Set-Variable -Name MyAddSoftwareFolder -Value ".ADDSOFTWARE" #-Option ReadOnly 
Set-Variable -Name MyHPIAActivityLogFile -Value 'activity.log' #-Option ReadOnly 
$s_AddAccesories = $false # FUTURE USE #

'My Public IP: '+(Invoke-WebRequest ifconfig.me/ip).Content.Trim() | Out-Host

#--------------------------------------------------------------------------------------
# error codes for color coding, etc.
Set-Variable -Name TypeError -Value -1 -Force
Set-Variable -Name TypeNorm -Value 1 -Force
Set-Variable -Name TypeWarn -Value 2 -Force
Set-Variable -Name TypeDebug -Value 4 -Force
Set-Variable -Name TypeSuccess -Value 5 -Force
Set-Variable -Name TypeNoNewline -Value 10 -Force

#=====================================================================================
#region: CMTraceLog Function formats logging in CMTrace style
function CMTraceLog {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $false)] $Message,
		[Parameter(Mandatory = $false)] $ErrorMessage,
		#[Parameter(Mandatory = $false)] $Component = "HP HPIA Repository Downloader",
		[Parameter(Mandatory = $false)] [int]$Type
	)
	<#
    Type: 1 = Normal, 2 = Warning (yellow), 3 = Error (red)
    #>
	$Time = Get-Date -Format "HH:mm:ss.ffffff"
	$Date = Get-Date -Format "MM-dd-yyyy"

	if ($null -ne $ErrorMessage) { $Type = $TypeError }
	#if ($Component -eq $null) { $Component = " " }
	if ($null -eq $Type) { $Type = $TypeNorm }

	#$LogMessage = "<![LOG[$Message $ErrorMessage" + "]LOG]!><time=`"$Time`" date=`"$Date`" component=`"$Component`" context=`"`" type=`"$Type`" thread=`"`" file=`"`">"
    $LogMessage = "<![LOG[$Message $ErrorMessage" + "]LOG]!><time=`"$Time`" date=`"$Date`" type=`"$Type`">"

    #$Type = 4: Debug output ($TypeDebug)
    #$Type = 10: no \newline ($TypeNoNewline)

    if ( ($Type -ne $TypeDebug) -or ( ($Type -eq $TypeDebug) -and $v_DebugMode) ) {
        $LogMessage | Out-File -Append -Encoding UTF8 -FilePath $Script:v_LogFile        
        OutToForm $Message $Type $Script:TextBox                        # if GUI is running, output to the form's textbox, else output to console        
    } else {
        $lineNum = ((get-pscallstack)[0].Location -split " line ")[1]   # output: CM_HPIARepo_Downloader.ps1: line 557
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
        $pMessage = '... {dbg}'+$pMessage
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

    if ( $v_DebugMode ) { CMTraceLog -Message "... Checking for required HP CMSL modules... " -Type $TypeNoNewline }

    # If module is imported say that and do nothing
    if (Get-Module | Where-Object {$_.Name -eq $m}) {
        if ( $v_DebugMode ) { write-host "Module $m is already imported." ; CMTraceLog -Message "Module already imported." -Type $TypSuccess }
    } else {

        # If module is not imported, but available on disk then import
        if (Get-Module -ListAvailable | Where-Object {$_.Name -eq $m}) {
            if ( $v_DebugMode ) { write-host "Importing Module $m." ; CMTraceLog -Message "Importing Module $m." -Type $TypeNoNewline}
            Import-Module $m -Verbose
            if ( $v_DebugMode ) { CMTraceLog -Message "Done" -Type $TypSuccess }
        } else {

            # If module is not imported, not available on disk, but is in online gallery then install and import
            if (Find-Module -Name $m | Where-Object {$_.Name -eq $m}) {
                if ( $v_DebugMode ) { write-host "Upgrading NuGet and updating PowerShellGet first." ; CMTraceLog -Message "Upgrading NuGet and updating PowerShellGet first." -Type $TypSuccess}
                if ( !(Get-packageprovider -Name nuget -Force) ) {
                    Install-PackageProvider -Name NuGet -ForceBootstrap
                }
                # next is PowerShellGet
                $lh_PSGet = find-module powershellget
                write-host 'Installing PowerShellGet version' $lh_PSGet.Version
                Install-Module -Name PowerShellGet -Force # -Verbose
                Write-Host 'should restart PowerShell after upating module PoweShellGet'

                # and finally module HPCMSL
                if ( $v_DebugMode ) { write-host "Installing and Importing Module $m." ; CMTraceLog -Message "Installing and Importing Module $m." -Type $TypSuccess}
                Install-Module -Name $m -Force -SkipPublisherCheck -AcceptLicense -Scope CurrentUser #  -Verbose 
                Import-Module $m -Verbose
                if ( $v_DebugMode ) { CMTraceLog -Message "Done" -Type $TypSuccess }
            } else {
                # If module is not imported, not available and not in online gallery then abort
                write-host "Module $m not imported, not available, and not in online gallery, exiting."
                if ( $v_DebugMode ) { CMTraceLog -Message "Module $m not imported, not available, and not in online gallery, exiting." -Type $TypError }
                exit 1
            }
        } # else if (Get-Module -ListAvailable | Where-Object {$_.Name -eq $m}) 
    } # else if (Get-Module | Where-Object {$_.Name -eq $m})

    # report CMSL version in use
    $hpcmsl = get-module -listavailable | Where-Object { $_.name -like 'HPCMSL' }
    #$HPCMSLVer = [string]$hpcmsl.Version.Major+'.'+[string]$hpcmsl.Version.Minor+'.'+[string]$hpcmsl.Version.Build
    CMTraceLog -Message "Using CMSL version: $($hpcmsl.Version)" -Type $TypError    

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
    $ga_dotRepository = "$($pRepoFolder)\.Repository"
    $ga_LastSyncEntry = $null
    $ga_CurrRepoRow = 0
    $ga_LastSyncLine = 0

    if ( Test-Path $ga_dotRepository ) {

        #--------------------------------------------------------------------------------
        # ga_SyncStartedString the last 'Sync started' entry 
        #--------------------------------------------------------------------------------
        $ga_ActivityLogFile = "$($ga_dotRepository)\$($MyHPIAActivityLogFile)"

        if ( Test-Path $ga_ActivityLogFile ) {
            $ga_SyncStartedString = 'sync has started'                     # look for this string in HPIA's log
            (Get-Content $ga_ActivityLogFile) | 
                Foreach-Object { 
                    $ga_CurrRepoRow++ 
                    if ($_ -match $ga_SyncStartedString) { $ga_LastSyncEntry = $_ ; $ga_LastSyncLine = $ga_CurrRepoRow } 
                } # Foreach-Object
            CMTraceLog -Message "      [activity.log - last sync @line $ga_LastSyncLine] $ga_LastSyncEntry " -Type $TypeWarn            
        } # if ( Test-Path $ga_ActivityLogFile )

        #--------------------------------------------------------------------------------
        # now, $ga_LastSyncLine holds the log's line # where the last sync started
        #--------------------------------------------------------------------------------
        if ( $ga_LastSyncEntry ) {

            $ga_LogFileContent = Get-Content $ga_ActivityLogFile

            for ( $i = 0; $i -lt $ga_LogFileContent.Count; $i++ ) {
                if ( $i -ge ($ga_LastSyncLine-1) ) {
                    if ( ($ga_LogFileContent[$i] -match 'done downloading exe') -or 
                        ($ga_LogFileContent[$i] -match 'already exists') ) {

                        CMTraceLog -Message "      [activity.log - Softpaq update] $($ga_LogFileContent[$i])" -Type $TypeWarn
                    }
                    if ( ($ga_LogFileContent[$i] -match 'Repository Synchronization failed') -or 
                        ($ga_LogFileContent[$i] -match 'ErrorRecord') -or
                        ($ga_LogFileContent[$i] -match 'WebException') ) {

                        CMTraceLog -Message "      [activity.log - Sync update] $($ga_LogFileContent[$i])" -Type $TypeError
                    } # if ( ($ga_LogFileContent[$i] -match 'Repository Synchronization failed') -or ...
                } # if ( $i -ge ($ga_LastSyncLine-1) )
            } # for ( $i = 0; $i -lt $ga_LogFileContent.Count; $i++ )
        } # if ( $ga_LastSyncEntry )

    } # if ( Test-Path $ga_dotRepository )

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
    
    if ( $v_DebugMode ) { CMTraceLog -Message "> init_repository()" -Type $TypeNorm }
    $ir_CurrentLoc = Get-Location # Save the current location to return later

    # Check if the provided repository folder path is valid and exists
    # Create the repository folder if it does not exist
    if ( -not (Test-Path $pRepoFolder) ) {    
        Try {
            New-Item -Path $pRepoFolder -ItemType directory | Out-Null
            CMTraceLog -Message "... repository path was not found, created: $pRepoFolder" -Type $TypeNorm
        } Catch {
            CMTraceLog -Message "... problem: $($_)" -Type $TypeError
        }
    } # if ( -not (Test-Path $pRepoFolder) )

    # resolve whether we need to initialze the repository or is already initialized
    switch ( $pInitialize ) {
        $true {
            Set-Location $pRepoFolder # Change to the repository folder for initialization
            Try {
                # Attempt to get repository info to check if it's already initialized
                Get-RepositoryInfo -ErrorAction Stop | Out-Null
                if ($v_DebugMode) { CMTraceLog -Message "... repository already initialized" -Type $TypeNorm }
            } Catch {
                # Catch block will handle the case where Get-RepositoryInfo fails, indicating it's not initialized               
                (Initialize-Repository) 6>&1
                if ( $v_DebugMode ) { CMTraceLog -Message  "... repository initialized"  -Type $TypeNorm }
                Set-RepositoryConfiguration -setting OfflineCacheMode -cachevalue Enable 6>&1 
                Set-RepositoryConfiguration -setting RepositoryReport -Format csv 6>&1
                if ( $v_DebugMode ) { CMTraceLog -Message  "... repository configured for HP Image Assistant" -Type $TypeNorm }
            } # Try/Catch block for Get-RepositoryInfo

            # next, intialize .ADDSOFTWARE folder for holding named softpaqs
            # This is where we will store add-on softpaqs for the repository
            # Check if the .ADDSOFTWARE folder exists, if not create it
            $ir_AddSoftpaqsFolder = "$($pRepoFolder)\$($MyAddSoftwareFolder)"
            if ( -not (Test-Path $ir_AddSoftpaqsFolder) ) {
                if ( $v_DebugMode ) { CMTraceLog -Message "... Adding `'Add-Ons`' Softpaqs folder $ir_AddSoftpaqsFolder" -Type $TypeNorm }
                New-Item -Path $ir_AddSoftpaqsFolder -ItemType directory | Out-Null
            } # if ( !(Test-Path $ir_AddSoftpaqsFolder) )
        } # $true
        $false {
            if ( $v_DebugMode ) { CMTraceLog -Message  "... Folder $($pRepoFolder) available" -Type $TypeNorm }
        } # $false
    } # switch ( $pInitialize )

    if ( $v_DebugMode ) { CMTraceLog -Message "< init_repository()" -Type $TypeNorm }

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
            (Get-Softpaq $pSoftpaq.Id) 6>&1
            CMTraceLog -Message "... done" -Type $TypeNorm
        } Catch {
            CMTraceLog -Message "... failed to download: $($_)" -Type $TypeError
        }          
    } # else if ( Test-Path $da_SoftpaqExePath ) {
    
    # download corresponding CVA file
    CMTraceLog -Message "`tdownloading $($pSoftpaq.id) CVA file" -Type $TypeNoNewline
    Try {
        Get-SoftpaqMetadataFile $pSoftpaq.Id -Overwrite 'Yes' 6>&1   # ALWAYS update the CVA file, even if previously downloaded
        CMTraceLog -Message "... done" -Type $TypeNorm    
    } Catch {
        CMTraceLog -Message "... failed to download: $($_)" -Type $TypeError
    }
    
    # download corresponding HTML file
    CMTraceLog -Message "`tdownloading $($pSoftpaq.id) HTML file" -Type $TypeNoNewline
    Try {
        $da_SoftpaqHtml = $pAddOnFolderPath+'\'+$pSoftpaq.id+'.html' # where to download to
        Invoke-WebRequest -UseBasicParsing -Uri $pSoftpaq.ReleaseNotes -OutFile "$da_SoftpaqHtml"
        CMTraceLog -Message "... done" -Type $TypeNorm
    } Catch {
        CMTraceLog -Message "... failed to download" -Type $TypeError
    }    
    
    Set-Location $da_CurrLocation
} # Function Download_Softpaq

Function Get_AddOnSoftpaqs {
[CmdletBinding()] 
    param( $pAddOnFlagFileFullPath )

    $gs_ProdCode = Split-Path $pAddOnFlagFileFullPath -leaf
    $gs_AddSoftpaqsFolder = Split-Path $pAddOnFlagFileFullPath -Parent

    [array]$gs_AddOnsList = Get-Content $pAddOnFlagFileFullPath

    if ( $gs_AddOnsList.count -ge 1 ) {
        if ( $v_DebugMode ) { CMTraceLog -Message "... platform $($gs_ProdCode): checking for Softpaq AddOns flag file" }
        
        Try {
            if ( $v_DebugMode ) { CMTraceLog -Message '...  > Get-SoftpaqList():'+$gs_ProdCode -Type $TypeNorm }
            $gs_SoftpaqList = (Get-SoftpaqList -platform $gs_ProdCode -os $Script:v_OS -osver $Script:v_OSVER -characteristic SSM) 6>&1  
            # check every reference file Softpaq for a match in the flg file list
            ForEach ( $gs_iEntry in $gs_AddOnsList ) {
                $gs_EntryFound = $False
                if ( $v_DebugMode ) { CMTraceLog -Message "...   Checking AddOn Softpaq by name: $($gs_iEntry)" -Type $TypeNorm }
                if ( $v_DebugMode ) { CMTraceLog -Message "...   # entries in ref file: $($gs_iEntry.Count)" -Type $TypeNorm }
                ForEach ( $gs_iSoftpaq in $gs_SoftpaqList ) {
                    if ( [string]$gs_iSoftpaq.Name -match $gs_iEntry ) {
                        $gs_EntryFound = $True
                        if ( $v_DebugMode ) { CMTraceLog -Message "... > Download_Softpaq(): $($gs_iEntry)" }
                        Download_Softpaq $gs_iSoftpaq $gs_AddSoftpaqsFolder
                    } # if ( [string]$gs_iSoftpaq.Name -match $gs_iEntry )
                } # ForEach ( $gs_iSoftpaq in $gs_SoftpaqList )
                if ( -not $gs_EntryFound ) {
                    if ( $v_DebugMode ) { CMTraceLog -Message  "... '$($gs_iEntry)': Softpaq not found for this platform and OS version"  -Type $TypeWarn }
                } # if ( -not $gs_EntryFound )
            } # ForEach ( $lEntry in $gs_AddOnsList )
        } Catch {
            if ( $v_DebugMode ) { CMTraceLog -Message "... $($gs_ProdCode): Error retrieving Reference file" -Type $TypeError }
        }                        
    } else {
        if ( $v_DebugMode ) { CMTraceLog -Message '    '$gs_ProdCode': Flag file found but empty ' }
    } # else if ( $gs_AddOnsList.count -ge 1 )

} # Function Get_AddOnSoftpaqs

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

    $da_AddSoftpaqsFolder = "$($pFolder)\$($MyAddSoftwareFolder)"   # get path of .ADDSOFTWARE subfolder
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
                #$da_ModelName = $pGrid[2,$iRow].Value                     # column 2 = Model name
                $da_AddOnsFlag = $pGrid.Rows[$iRow].Cells['AddOns'].Value # column 7 = 'AddOns' checkmark
                $da_ProdIDFlagFile = $da_AddSoftpaqsFolder+'\'+$da_ProdCode

                # if user checked the addOns column... and the flag file is there...
                if ( $da_AddOnsFlag -and (Test-Path $da_ProdIDFlagFile) ) {                    
                    Get_AddOnSoftpaqs $da_ProdIDFlagFile 
                } # if (Test-Path $da_ProdIDFlagFile)

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
                    $da_ProdIDFlagFile = $da_AddSoftpaqsFolder+'\'+$da_SystemID
                    if ( $da_AddOnsValue -and (Test-Path $da_ProdIDFlagFile) ) {                   
                        if ( $v_DebugMode ) { CMTraceLog -Message 'Flag file found:'+$da_ProdIDFlagFile -Type $TypeWarn }
                        Get_AddOnSoftpaqs $da_ProdIDFlagFile                   
                    } # if ( $da_AddOnsValue -and (Test-Path $da_ProdIDFlagFile) )                    
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
        if ( $v_DebugMode ) { CMTraceLog -Message "...    Restoring named softpaqs/cva/html files to Repository" -Type $TypeNoNewline }
        $ra_source = "$($pFolder)\$($MyAddSoftwareFolder)\*.*"
        Copy-Item -Path $ra_source -Filter *.cva -Destination $pFolder
        Copy-Item -Path $ra_source -Filter *.exe -Destination $pFolder
        Copy-Item -Path $ra_source -Filter *.html -Destination $pFolder
        if ( $v_DebugMode ) { CMTraceLog -Message " - completed" -Type TypeNorm }
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
            if ( $v_DebugMode ) { CMTraceLog -Message  '...   calling CMSL Invoke-RepositorySync' -Type $TypeNorm }
            
            Invoke-RepositorySync 6>&1
                
            # find what sync'd from the CMSL log file for this run
            try {
                if ( $v_DebugMode )  { CMTraceLog -Message  '...   retrieving repository filters' -Type $TypeNorm }
                #$lRepoFilters = (Get-RepositoryInfo).Filters  # see if any filters are used to continue

                Get_ActivityLogEntries $pFolder  # get sync'd info from HPIA activity log file
                $sc_contentsFile = "$($pFolder)\.repository\contents.csv"
                if ( Test-Path $sc_contentsFile ) {
                    $scr_ContentsHASH = (Get-FileHash -Path "$($pFolder)\.repository\contents.csv" -Algorithm SHA256).Hash
                    if ( $v_DebugMode ) { CMTraceLog -Message "... SHA256 Hash of 'contents.csv': $($scr_ContentsHASH)" -Type $TypeWarn }
                }              
                #--------------------------------------------------------------------------------
                if ( $v_DebugMode ) { CMTraceLog -Message  '...   calling CMSL invoke-RepositoryCleanup' -Type $TypeNorm }
                Invoke-RepositoryCleanup 6>&1

                if ( $v_DebugMode ) { CMTraceLog -Message  '...   > Download_AddOns()' -Type $TypeNorm }
                Download_AddOns $pGrid $pFolder $pCommonFlag
                if ( $v_DebugMode )  {CMTraceLog -Message  '...   > Restore_Softpaqs()' -Type $TypeNorm }
                Restore_AddOns $pFolder (-not $script:noIniSw)
                #--------------------------------------------------------------------------------
                # see if Cleanup modified the contents.csv file 
                # - seems like (up to at least 1.6.3) RepositoryCleanup does not Modify 'contents.csv'
                if ( Test-Path $sc_contentsFile ) {
                    $scr_ContentsHASH = Get_SyncdContents $pFolder       # Sync command creates 'contents.csv'
                    if ( $v_DebugMode )  { CMTraceLog -Message "...    MD5 Hash of 'contents.csv': $($scr_ContentsHASH)" -Type $TypeWarn }
                }                
            } catch {
                $scr_ContentsHASH = $null
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
    foreach ( $cat in $Script:v_FilterCategories ) {
        # if the checkbox for this category is checked (true), then add the filter
        if ( $pGrid.Rows[$pRow].Cells[$cat].Value ) {
            if ( $v_DebugMode )  { CMTraceLog -Message  "... adding filter: -Platform $pModelID -os $Script:v_OS:$Script:v_OSVER -category $($cat) -characteristic ssm" -Type $TypeNoNewline }
            $umf_Res = (Add-RepositoryFilter -platform $pModelID -os $Script:v_OS -osver $Script:v_OSVER -category $cat -characteristic ssm 6>&1)          
            if ( $v_DebugMode ) { CMTraceLog -Message $umf_Res -Type $TypeWarn }
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

    $sr_CurrentSetLoc = Get-Location

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
            $sr_ProdFolders = Get-ChildItem -Path $Script:v_Root_IndividualRepoFolder | Where-Object {($_.psiscontainer)}

            foreach ( $iprodName in $sr_ProdFolders ) {
                $sr_CurrentPath = "$($Script:v_Root_IndividualRepoFolder)\$($iprodName.name)"
                set-location $sr_CurrentPath
                Sync_and_Cleanup_Repository $pGrid $sr_CurrentPath $pCommonRepoFlag
            } # foreach ( $iprodName in $sr_ProdFolders )
        } else {
            CMTraceLog -Message  "... Shared/Individual Repository Folder selected, Head repository not initialized" -Type $TypeNorm
        } # else if ( !(Test-Path $Script:v_Root_IndividualRepoFolder) ) 
    } # else if ( $Script:v_CommonRepo )

    CMTraceLog -Message  "Sync DONE" -Type $TypeSuccess
    Set-Location -Path $sr_CurrentSetLoc

} # Sync_Repos

#=====================================================================================
<#
    Function sync_individualRepositories
        for every selected model, go through every repository by model
            - ensure there is a valid repository
            - remove all filters unless Keep OS Filters is set to $true (this will keep existing filters in place)
            - add filters based on the selected categories in the UI for that model
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
    if ( $v_DebugMode ) { CMTraceLog -Message "... at folder: $($pRepoHeadFolder)" -Type $TypeDebug }
    if ( $v_DebugMode ) { CMTraceLog -Message "... for (checked item) rows: $($pCheckedItemsList)" -Type $TypeDebug }

    #--------------------------------------------------------------------------------
    # decide if we work on single repo or individual repos
    #--------------------------------------------------------------------------------     
    init_repository $pRepoHeadFolder $false             # make sure Main repo folder exists, do NOT make it a repository

    #--------------------------------------------------------------------------------
    if ( $v_DebugMode ) { CMTraceLog -Message "... stepping through all selected models" -Type $TypeDebug }

    # go through every selected HP Model in the list
    for ( $i = 0; $i -lt $pGrid.RowCount; $i++ ) {
        $si_ModelId = $pGrid[1,$i].Value                    # column 1 has the Model/Prod ID
        $si_ModelName = $pGrid[2,$i].Value                  # column 2 has the Model name

        # if model entry is checked, we need to create a repository, but ONLY if $Script:v_CommonRepo = $false
        if ( $i -in $pCheckedItemsList ) {

            $si_RepoFolder = "$($pRepoHeadFolder)\$($si_ModelId)_$($si_ModelName)" # this is the repository folder for this model, with SysID in name
            if ( -not (Test-Path -Path $si_RepoFolder) ) {
                $si_RepoFolder = "$($pRepoHeadFolder)\$($si_ModelName)"         # this is the repo folder for this model, without SysID in name
            }
            if ( $v_DebugMode ) { CMTraceLog -Message "`n... Updating model: $si_ModelId : $si_ModelName -- path "+$si_RepoFolder -Type $TypeDebug }
            
            init_repository $si_RepoFolder $true            # initialize the individual repository folder for this model, if it does not exist   
            
            Update_Model_Filters $pGrid $si_RepoFolder $si_ModelId $i           # update filters for the current model in the grid
            
            Sync_and_Cleanup_Repository $pGrid $si_RepoFolder $False            # invoke sync and cleanup for this individual repository folder

            # update SCCM Repository package, if user checked that off
            if ( $Script:v_UpdateCMPackages ) {
                CM_RepoUpdate $si_ModelName $si_ModelId $si_RepoFolder
            }
            
        } # if ( $i -in $pCheckedItemsList )

    } # for ( $i = 0; $i -lt $pGrid.RowCount; $i++ )

    if ( $pNewModels ) { Update_INIModelsFromGrid $pGrid $pCommonRepoFolder }

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

    $sc_CurrentLoc = Get-Location
    $sc_UnselectedRows = $false # initialize to check for unselected rows in the grid

    CMTraceLog -Message "> Sync_CommonRepository - START" -Type $TypeNorm
    if ( $v_DebugMode ) { CMTraceLog -Message "... for OS version: $($Script:v_OSVER)" -Type $TypeDebug }

    init_repository $pRepoFolder $true              # make sure repo folder exists, or create it, initialize

    if ( $Script:v_KeepFilters ) {
        CMTraceLog -Message  "... Keeping existing filters" -Type $TypeNorm
    } else {
        CMTraceLog -Message  "... Removing existing filters" -Type $TypeNorm
    } # else if ( $Script:v_KeepFilters )

    # search for unselected rows in the grid to determine if any models were unchecked
    # this will help determine if we should ask the user to confirm Sync, if there are unselected rows
    for ( $iRow = 0; $iRow -lt $pGrid.RowCount; $iRow++ ) { if ( $pGrid[0,$iRow].Value -eq $false ) {  $sc_UnselectedRows = $true ; break } }

    $sc_SyncAsk = 'Yes'                         # default to 'Yes' to proceed with sync even if unselected rows found

    if ( $sc_UnselectedRows ) {
        # if there were unselected rows, log it
        CMTraceLog -Message "... some models were not selected in the grid, filters will not be applied to those models" -Type $TypeWarn
        $sc_SyncAsk = AskUser 'There are unselected entries. OK to Sync and remove their content?' 4 # 4="YesNo"
    } # if ( $sc_UnselectedRows )

    if ( $sc_SyncAsk -eq 'Yes') {
        # now proceed with the sync operation for the common repository, one row at a time
        for ( $iRow = $pGrid.RowCount; $iRow -gt 0; $iRow-- ) {
            $sc_GridRow = $iRow - 1 # adjust for 0-based index, since we are looping downwards to avoid index issues when removing rows

            #$sc_ModelSelected = $pGrid[0,$iRow].Value           
            $sc_ModelId = $pGrid[1,$sc_GridRow].Value                 # col 1 has the SysID
            $sc_ModelName = $pGrid[2,$sc_GridRow].Value               # col 2 has the Model name
            #$sc_AddOnsFlag = $pGrid.Rows[$iRow].Cells['AddOns'].Value # column 7 has the AddOns checkmark

            switch ( $pGrid[0,$sc_GridRow].Value ) {
                $true   {
                    # if the model is selected, but we are keeping filters, then just update the filters for the model
                    if ( -not $Script:v_KeepFilters ) {
                        if ( ((Get-RepositoryInfo).Filters).count -gt 0 ) { Remove-RepositoryFilter -platform $sc_ModelId -yes 6>&1 }
                        if ( $v_DebugMode ) { CMTraceLog -Message "... removed existing filters for platform: $($sc_ModelId):$($sc_ModelName)" -Type $TypeWarn }
                    }
                    Update_Model_Filters $pGrid $pRepoFolder $sc_ModelId $sc_GridRow # update filters for the current model in the grid
                } # $true
                
                # # if the model is not selected and we are keeping filters, then do not remove filters, just update the repository path
                # this is to ensure that if a model is selected but filters are kept, it will not remove the filters for the selected model
                $false  {
                    if ( -not $Script:v_KeepFilters ) {
                        if ( ((Get-RepositoryInfo).Filters).count -gt 0 ) { Remove-RepositoryFilter -platform $sc_ModelId -yes 6>&1 }
                        if ( $v_DebugMode ) { CMTraceLog -Message "... removed existing filters for platform: $($sc_ModelId):$($sc_ModelName)" -Type $TypeWarn }
                        Update_Model_Filters $pGrid $pRepoFolder $sc_ModelId $sc_GridRow # update filters for the current model in the grid
                        Remove_SinglePlatform $pGrid $sc_GridRow $pRepoFolder $true
                    } # if ( -not $Script:v_KeepFilters )                 
                } # $false
            } # switch ( $pGrid[0,$iRow].Value )
           
            #Update_Model_Filters $pGrid $pRepoFolder $sc_ModelId $iRow # update filters for the current model in the grid
        } # for ( $iRow = 0; $iRow -lt $pGrid.RowCount; $iRow++ )

        Update_INIModelsFromGrid $pGrid $pRepoFolder    # update the INI file with the current models in the grid, if needed
        Sync_and_Cleanup_Repository $pGrid $pRepoFolder $True # $True = sync a common repository

    } # if ( $sc_SyncAsk -eq 'Yes')
    
    Set-Location -Path $sc_CurrentLoc

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

    $lf_CurrentSetLoc = Get-Location

    if ( $Script:v_CommonRepo ) {
        # basic check to confirm the repository was configured for HPIA
        if ( !(Test-Path "$($Script:v_Root_CommonRepoFolder)\.Repository") ) {
            CMTraceLog -Message "... Common Repository Folder selected, not initialized" -Type $TypeNorm 
            "... Common Repository Folder selected, not initialized for HPIA"
            return
        } 
        set-location $Script:v_Root_CommonRepoFolder

        $lf_ProdFilters = (get-repositoryinfo).Filters
        $lf_platforms = $lf_ProdFilters.platform | Get-Unique

        # develop the filter list by product
        foreach ( $iProduct in $lf_platforms ) {
            $lf_characteristics = $lf_ProdFilters.characteristic | Get-Unique 
            $lf_OS = $lf_ProdFilters.operatingSystem | Get-Unique
            $lf_Msg = "   -Platform $($iProduct) -OS $($lf_OS) -Category"
            foreach ( $iFilter in $lf_ProdFilters ) { if ( $iFilter.Platform -eq $iProduct ) { $lf_Msg += ' '+$iFilter.category } }
            $lf_Msg += " -Characteristic $($lf_characteristics)"
            CMTraceLog -Message $lf_Msg -Type $TypeNorm
        }
    } else {
        # basic check to confirm the head repository exists
        if ( !(Test-Path $Script:v_Root_IndividualRepoFolder) ) {
            CMTraceLog -Message "... Shared/Individual Repository Folder selected, requested folder NOT found" -Type $TypeNorm
            "... Shared/Individual Repository Folder selected, Head repository not initialized"
            return
        } # if ( !(Test-Path $Script:v_Root_IndividualRepoFolder) )
        set-location $Script:v_Root_IndividualRepoFolder | Where-Object {($_.psiscontainer)}

        # let's traverse every product Repository folder
        $lf_ProdFolders = Get-ChildItem -Path $Script:v_Root_IndividualRepoFolder

        foreach ( $iprodName in $lf_ProdFolders ) {
            set-location "$($Script:v_Root_IndividualRepoFolder)\$($iprodName.name)"

            $lf_ProdFilters = (get-repositoryinfo).Filters
            $lf_platforms = $lf_ProdFilters.platform | Get-Unique 
            $lf_characteristics = $lf_ProdFilters.characteristic | Get-Unique            
            $lf_OS = $lf_ProdFilters.operatingSystem | Get-Unique 

            $lf_Msg = "   -Platform $($lf_platforms) -OS $($lf_OS) -Category"
            foreach ( $icat in $lf_ProdFilters.category ) { $lf_Msg += ' '+$icat }
            $lf_Msg += " -Characteristic $($lf_characteristics)"
            CMTraceLog -Message $lf_Msg -Type $TypeNorm
        } # foreach ( $lprodName in $lf_ProdFolders )

    } # else if ( $Script:v_CommonRepo )

    Set-Location -Path $lf_CurrentSetLoc

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
        if ( $v_DebugMode ) {CMTraceLog -Message "... Refreshing Grid from Common Repository: ''$pCommonFolder'' (calling clear_grid())" -Type $TypeNorm }
        clear_grid $pGrid
    } else {
        if ( $v_DebugMode ) { CMTraceLog -Message "... Filters from Common Repository ...$($pCommonFolder)" -Type $TypeNorm }
    }
    #---------------------------------------------------------------
    # Let's now get all the HPIA filters from the repository
    #---------------------------------------------------------------
    set-location $pCommonFolder
    $gc_ProdFilters = (get-repositoryinfo).Filters   # get the list of filters from the repository

    ### example filter returned by 'get-repositoryinfo' CMSL command: 
    ###
    ### platform        : 8438
    ### operatingSystem : win10:2004 win10:2004
    ### category        : BIOS firmware
    ### releaseType     : *
    ### characteristic  : ssm

    foreach ( $filter in $gc_ProdFilters ) {
        # check each row SysID against the Filter Platform ID
        for ( $i = 0; $i -lt $pGrid.RowCount; $i++ ) {
            $gc_Platform = $filter.platform

            if ( $pRefreshGrid ) {
                if ( $gc_Platform -eq $pGrid[1,$i].value ) {
                    # we matched the row/SysId with the Filter

                    # ... so let's add each category in the filter to the Model in the GUI
                    foreach ( $cat in  ($filter.category.split(' ')) ) {
                        $pGrid.Rows[$i].Cells[$cat].Value = $true
                    }
                    $pGrid[0,$i].Value = $true   # check the selection column 0
                    $pGrid[($pGrid.ColumnCount-1),$i].Value = $pCommonFolder # ... and the Repository Path
                
                    # if we are maintaining AddOns Softpaqs, show it in the checkbox
                    $gc_AddOnsRepoFile = $pCommonFolder+'\'+$MyAddSoftwareFolder+'\'+$gc_Platform 

                    if ( (Test-Path $gc_AddOnsRepoFile) -and (-not ($pGrid.Rows[$i].Cells['AddOns'].Value))) {
                        $pGrid.Rows[$i].Cells['AddOns'].Value = $True
                        [array]$gc_AddOns = Get-Content $gc_AddOnsRepoFile
                        if ( $gc_AddOns.count -gt 0 ) {
                            if ( $v_DebugMode ) { CMTraceLog -Message "...  Additional Softpaqs to be maintained for '$gc_Platform': { $gc_AddOns }" -Type $TypeWarn }
                        } # if ( $gc_AddOns.count -gt 0 )
                    }
                } # if ( $gc_Platform -eq $pGrid[1,$i].value )
            } else {
                if ( $v_DebugMode ) { CMTraceLog -Message "... listing filters (calling List_Filters())" -Type $TypeWarn }
                List_Filters
            } # else if ( $pRefreshGrid )
        } # for ( $i = 0; $i -lt $pGrid.RowCount; $i++ )
    } # foreach ( $filter in $gc_ProdFilters )

    [void]$pGrid.Refresh
    if ( $v_DebugMode ) { CMTraceLog -Message "... Refreshing Grid ...DONE" -Type $TypeSuccess }
    if ( $v_DebugMode ) { CMTraceLog -Message '< Get_CommonRepofilters()' -Type $TypeNorm }
    
} # Function Get_CommonRepofilters

#=====================================================================================
<#
    Function Get_IndividualRepofilters
        Retrieves category filters from the repository for each model from the 
        ... and populates the Grid appropriately
        Parameters:
            $pGrid                          The models grid in the GUI
            $pRepoLocation                      Where to start looking for repositories
#>
Function Get_IndividualRepofilters {
    [CmdletBinding()]
	param( $pGrid,                                  # array of row lines that are checked
        $pRepoRoot )

    if ( $v_DebugMode ) { CMTraceLog -Message '   > Get_IndividualRepofilters() START' -Type $TypeNorm }

    set-location $pRepoRoot

    if ( $v_DebugMode ) { CMTraceLog -Message '...  Refreshing Grid from Individual Repositories ...' -Type $TypeNorm }

    #--------------------------------------------------------------------------------
    # now check for each product's repository folder
    # if the repo is created, then check the category filters
    #--------------------------------------------------------------------------------
    for ( $iRow = 0; $iRow -lt $pGrid.RowCount; $iRow++ ) {

        $gi_ModelId = $pGrid[1,$iRow].Value                                     # column 1 has the Model/Prod ID
        $gi_ModelName = $pGrid[2,$iRow].Value                                   # column 2 has the Model name 
        
        $gi_PlatformRepoFolder = "$($pRepoRoot)\$($gi_ModelId)_$($gi_ModelName)"  # this is the repo folder - SysID in the folder name
        if ( -not (Test-Path -Path $gi_PlatformRepoFolder) ) {
            $gi_PlatformRepoFolder = "$($pRepoRoot)\$($gi_ModelName)"           # this is the repo folder for this model, without SysID in name
        }

        if ( Test-Path $gi_PlatformRepoFolder ) {
            set-location $gi_PlatformRepoFolder                                 # move to location of Repository
                                    
            $pGrid[($dataGridView.ColumnCount-1),$iRow].Value = $gi_PlatformRepoFolder # set the location of Repository in the grid row
            
            # if we are maintaining AddOns Softpaqs, show it in the checkbox
            $gi_AddOnRepoFile = $gi_PlatformRepoFolder+'\'+$MyAddSoftwareFolder+'\'+$gi_ModelId   
            $pGrid.Rows[$iRow].Cells['AddOns'].Value = $False                   # default: assume the AddOns flag file is not present or is empty
            [array]$gi_AddOnsList = Get-Content $gi_AddOnRepoFile
            if ( (Test-Path $gi_AddOnRepoFile) -and ($gi_AddOnsList.count -gt 0) ) {  
                $pGrid.Rows[$iRow].Cells['AddOns'].Value = $True                # the flag file exists and is not empty
                if ( $v_DebugMode ) { CMTraceLog -Message "...  Additional Softpaqs to be maintained for '$gi_ModelName': { $gi_AddOnsList }" -Type $TypeWarn }
            } # if ( Test-Path $gi_AddOnRepoFile )

            <# populate the grid with the filters for this model
                # Example output from get-repositoryinfo for filters:
                platform        : 8715
                operatingSystem : win10:21h1
                category        : BIOS
                releaseType     : *
                characteristic  : ssm
            #>
            $gi_ProdFilters = (get-repositoryinfo).Filters                      # get the list of filters from the repository
            foreach ( $iEntry in $gi_ProdFilters ) {
                foreach ( $cat in  ($gi_ProdFilters.category.split(' ')) ) {
                    $pGrid.Rows[$iRow].Cells[$cat].Value = $true
                } # foreach ( $cat in  ($gi_ProdFilters.category.split(' ')) )
            } # foreach ( $iEntry in $gi_ProdFilters )
        } # if ( Test-Path $gi_PlatformRepoFolder )
    } # for ( $iRow = 0; $iRow -lt $pModelsList.RowCount; $iRow++ ) 

    if ( $v_DebugMode ) { CMTraceLog -Message '   < Get_IndividualRepofilters() DONE' -Type $TypeNorm }
    
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

    if ( $v_DebugMode ) { CMTraceLog -Message ">> Browse Individual Repositories - Start" -Type $TypeSuccess }

    # ask the user for the repository

    $bi_browse = New-Object System.Windows.Forms.FolderBrowserDialog
    $bi_browse.SelectedPath = $pCurrentRepository         
    $bi_browse.Description = "Select a Root Folder for individual repositories"
    $bi_browse.ShowNewFolderButton = $true
                                  
    if ( $bi_browse.ShowDialog() -eq "OK" ) {
        $bi_Repository = $bi_browse.SelectedPath
        Import_IndividualRepos $pGrid $bi_browse.SelectedPath
    } else {
        $bi_Repository = $null
    } # else if ( $bi_browse.ShowDialog() -eq "OK" )
    
    $bi_browse.Dispose()

    if ( $v_DebugMode ) { CMTraceLog -Message "<< Browse Individual Repositories - Done" -Type $TypeSuccess }

    Return $bi_Repository

} # Function Browse_IndividualRepos

#=====================================================================================
<#
    Function Import_IndividualRepos
        - Populate list of Platforms from Individual Repositories
        - also, update INI file about imported repository models if user agrees
        Paramaters: $pGrid = pointer to model grid in UI
                    $pCurrentRepository = head of individual repository folders
                    
        returns: object w/Folder picked for repository, or $null if user cancelled
                Folder return value contains True/Fail entry [0] + folder name [1]
#>
Function Import_IndividualRepos {
    [CmdletBinding()]
	param( $pGrid, $pCurrentRepository )
    if ( $v_DebugMode ) { CMTraceLog -Message "> Import_IndividualRepos START - at ($($pCurrentRepository))" -Type $TypeNorm }

    if ( $v_DebugMode ) { CMTraceLog -Message  "... calling Empty_Grid() to clear list" -Type $TypeNorm }
    Empty_Grid $pGrid
    if ( -not (Test-Path $pCurrentRepository) ) {
        init_repository $pCurrentRepository $False  # $false means this is not a HPIA repo, so don't init as such
    }
    Set-Location $pCurrentRepository
    if ( $v_DebugMode ) { CMTraceLog -Message  "... checking for existing repositories" -Type $TypeNorm }
    $ir_Directories = (Get-ChildItem -Path $pCurrentRepository -Directory) | Where-Object { $_.BaseName -notmatch "_BK$" }
    
    $ir_ReposFound = $False     # assume no repositories exist here

    foreach ( $iFolder in $ir_Directories ) { 
        $ir_RepoFolder = $pCurrentRepository+'\'+$iFolder            
        Set-Location $ir_RepoFolder
        Try {
            CMTraceLog -Message  "... [Repository Found ($($ir_RepoFolder)) -- adding model to grid]" -Type $TypeNorm
            # obtain platform SysID from filters in the repository
            [array]$ir_RepoPlatforms = (Get-RepositoryInfo).Filters.platform | Get-Unique
            # add the platforms found in the repository to the grid
            $ir_RepoPlatforms | ForEach-Object {
                [void]$pGrid.Rows.Add(@( $true, $_, $iFolder ))                
            } # $ir_RepoPlatforms | foreach           
            $ir_ReposFound = $True
        } Catch {
            CMTraceLog -Message  "... $($ir_RepoFolder) is NOT a Repository" -Type $TypeNorm
        } # Catch
    } # foreach ( $iFolder in $ir_Directories )
    
    if ( $ir_ReposFound ) {
        if ( $v_DebugMode ) { CMTraceLog -Message  "... > Get_IndividualRepofilters() - retrieving Filters" }
        Get_IndividualRepofilters $pGrid $pCurrentRepository
        if ( $v_DebugMode ) { CMTraceLog -Message  "... > Update_INIModelsFromGrid()" }
        Update_INIModelsFromGrid $pGrid $pCurrentRepository  
    } else {
        if ( $v_DebugMode ) { CMTraceLog -Message  "... no repositories found" }
        #$bi_ask = [System.Windows.MessageBox]::Show('Clear the HPModel list in the INI file?','INI File update','YesNo')
        $ir_askClear = AskUser 'Clear the HPModel list in the INI file?' 4 # 4="YesNo"
        if ( $ir_askClear -eq 'Yes' ) {
            if ( $v_DebugMode ) { CMTraceLog -Message  "...calling Update_UIandINISetting() to clear HPModels list " }
            Update_UIandINISetting $pCurrentRepository $False              # $False = using individual repositories
        } else {
            CMTraceLog -Message  "... previous HPModels list remain in INI file"
        }
    } # else if ( $ir_ReposFound )

    if ( $v_DebugMode ) { CMTraceLog -Message '< Import_IndividualRepos DONE' -Type $TypeSuccess }
} # Function Import_IndividualRepos

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

    if ( $v_DebugMode ) { CMTraceLog -Message '> Browse_CommonRepo START' -Type $TypeNorm }
    # let's find the repository to import or use the current

    $bc_browse = New-Object System.Windows.Forms.FolderBrowserDialog
    $bc_browse.SelectedPath = $pCurrentRepository        # start with the repo listed in the INI file
    $bc_browse.Description = "Select an HPIA Common/Shared Repository"
    $bc_browse.ShowNewFolderButton = $true

    if ( $bc_browse.ShowDialog() -eq "OK" ) {
        $bc_Repository = $bc_browse.SelectedPath
        Import_Repository $pGrid $bc_Repository $pModelsTable    
        if ( $v_DebugMode ) { CMTraceLog -Message "< Browse Common Repository Done ($($bc_Repository))" -Type $TypeSuccess }
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

    if ( $v_DebugMode ) { CMTraceLog -Message "> Import_Repository: $($pRepoFolder)" -Type $TypeNorm }

    Empty_Grid $pGrid

    if ( [string]::IsNullOrEmpty($pRepoFolder ) ) {
        CMTraceLog -Message  "... There is no Repository to import" -Type $TypeWarn        
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
            if ( $v_DebugMode ) { CMTraceLog -Message  "... this is a valid HPIA Repository" -Type $TypeNorm }

            #--------------------------------------------------------------------------------
            #if ( $ir_ProdFilters.count -eq 0 -and ($ir_FlagFiles.count -eq 0) ) {
            if ( $ir_ProdFilters.count -eq 0 ) {
                if ( $v_DebugMode ) { CMTraceLog -Message  "... No repository filters found" -Type $TypeNorm }
                if ( $v_DebugMode ) { CMTraceLog -Message  "... Updating UI and INI file from filters - > Update_UIandINISetting()" -Type $TypeNorm }
                Update_UIandINISetting $pRepoFolder $True     # $True = common repo
                if ( $v_DebugMode ) { CMTraceLog -Message  "... Updating models in INI file - > Update_INIModelsFromGrid()" -Type $TypeNorm }
                Update_INIModelsFromGrid $pGrid $pRepoFolder  
            } else {
                #--------------------------------------------------------------------------------
                # fill the grid with the platforms from the repository filters if any exist 
                # ... first, get the (unique) list of platform SysIDs in the repository from the filters
                #--------------------------------------------------------------------------------
                [array]$ir_RepoPlatforms = (Get-RepositoryInfo).Filters.platform | Get-Unique
                # next, add each product to the grid, to then populate with the filters
                if ( $v_DebugMode ) { CMTraceLog -Message  "... Adding platforms to Grid from repository" }
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

                if ( $v_DebugMode ) { CMTraceLog -Message "... Finding filters from $($pRepoFolder) (calling Get_CommonRepofilters())" -Type $TypeNorm }
                Get_CommonRepofilters $pGrid $pRepoFolder $True

                if ( $v_DebugMode ) { CMTraceLog -Message  "... Updating UI and INI file from filters (calling Update_UIandINISetting())" -Type $TypeNorm }
                Update_UIandINISetting $pRepoFolder $True     # $True = common repo

                if ( $v_DebugMode ) { CMTraceLog -Message  "... Updating models in INI file (calling Update_INIModelsFromGrid())" -Type $TypeNorm }
                Update_INIModelsFromGrid $pGrid $pRepoFolder  
            } # if ( $ir_ProdFilters.count -gt 0 )
            if ( $v_DebugMode ) { CMTraceLog -Message "< Import Repository Done" -Type $TypeSuccess }
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

    if ( $v_DebugMode ) { CMTraceLog -Message "> Check_PlatformsOSVersion($($pOS)/$($pOSVersion))" -Type $TypeNorm }

    # search thru the table entries for checked items, and see if each product
    # has support for the selected OS version

    for ( $i = 0; $i -lt $pGrid.RowCount; $i++ ) {

        if ( $pGrid[0,$i].Value ) {
            $cp_Platform = $pGrid[1,$i].Value
            $cp_PlatformName = $pGrid[2,$i].Value

            # get list of OS versions supported by this platform and OS version
            $cp_OSList = get-hpdevicedetails -platform $cp_Platform -OSList -ErrorAction Continue

            $cp_OSIsSupported = $false
            foreach ( $entry in $cp_OSList ) {
                if ( ($pOSVersion -eq $entry.OperatingSystemRelease) -and $entry.OperatingSystem -match $pOS.substring(3) ) {
                    if ( $v_DebugMode ) { CMTraceLog -Message  "... $pOS/$pOSVersion OS is supported for $($cp_Platform)/$($cp_PlatformName)" -Type $TypeNorm }
                    $cp_OSIsSupported = $true
                }
            } # foreach ( $entry in $lOSList )
            if ( -not $cp_OSIsSupported ) {
                CMTraceLog -Message  "... $pOS/$pOSVersion OS is NOT supported for $($cp_Platform)/$($cp_PlatformName)" -Type $TypeWarn
            }
        } # if ( $dataGridView[0,$i].Value )  

    } # for ( $i = 0; $i -lt $pGrid.RowCount; $i++ )
    
    if ( $v_DebugMode ) { CMTraceLog -Message '< Check_PlatformsOSVersion()' -Type $TypeNorm }

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
    if ( $v_DebugMode ) { CMTraceLog -Message $pText -Type $TypeNorm }
    
} # Mod_INISetting

#=====================================================================================
<#
    Function Update_INIModelsFromGrid
        Create $HPModelsTable array from current checked grid entries and updates INI file with it
        Parameters
            $pGrid
            $pRepositoryFolder
#>
Function Update_INIModelsFromGrid {
    [CmdletBinding()]
	param( $pGrid,
        $pRepositoryFolder )            # required to find the platform ID flag files hosted in .ADDSOFTWARE

    if ( $v_DebugMode ) { CMTraceLog -Message '... > Update_INIModelsFromGrid()' -Type $TypeNorm }

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
    if ( $ui_ModelsList.count -gt 0 ) {
        if ( $v_DebugMode ) { CMTraceLog -Message "...     Updating INI file with these platforms:  $($ui_ModelsList)" -Type $TypeNorm }
    } # if ( $ui_ModelsList.count -gt 0 )
    
    # ---------------------------------------------------------
    # Now, replace HPModelTable in INI file with list from grid
    # ---------------------------------------------------------
    if ( Test-Path $ui_INIFilePath ) {
        if ( $v_DebugMode ) { CMTraceLog -Message "... Saving Models list to INI File $ui_INIFilePath" -Type $TypeNorm }
        $ui_ListHeader = '\$HPModelsTable = .*'
        $ui_ProdsEntry = '.*@{ ProdCode = .*'
        # remove existing model entries from file
        Set-Content -Path $ui_INIFilePath -Value (get-content -Path $ui_INIFilePath | Select-String -Pattern $ui_ProdsEntry -NotMatch)
        # add new model lines        
        (get-content $ui_INIFilePath) -replace $ui_ListHeader, "$&$($ui_ModelsList)" | Set-Content $ui_INIFilePath
    } else {
        CMTraceLog -Message " ... INI file was not found" -Type $TypeWarn
    } # else if ( Test-Path $ga_ActivityLogFile )

    if ( $v_DebugMode ) { CMTraceLog -Message '... < Update_INIModelsFromGrid()' -Type $TypeNorm }
    
} # Function Update_INIModelsFromGrid 

#=====================================================================================
<#
    Function Update_GridFromINI
    This is the MAIN function with a Gui that sets things up for the user
#>
Function Update_GridFromINI {
    [CmdletBinding()]
	param( $pGrid,
            $pModelsTable )

    CMTraceLog -Message "... populating Grid from INI file's `$HPModels list" -Type $TypeWarn
    <# example INI platform list
        @{ ProdCode = '8723'; Model = 'HP ZBook Firefly 14 G7 Mobile Workstation' }
        @{ ProdCode = '8ABB 896D'; Model = 'HP EliteBook 840 14 inch G9 Notebook PC' }
    #> 
    $pModelsTable | 
        ForEach-Object {
            # handle case of multiple ProdCodes in model entry
            foreach ( $iSysID in $_.ProdCode.split(' ') ) {
                [void]$pGrid.Rows.Add( @( $False, $iSysID, $_.Model) )   # populate checkmark, ProdId, Model Name
                if ( $v_DebugMode ) { CMTraceLog -Message "... adding $($iSysID):$($_.Model)" -Type $TypeNorm }
            } # foreach ( $iSysID in $_.ProdCode.split(' ') )
        } # ForEach-Object

    $pGrid.Refresh()

} # Update_GridFromINI

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
        $CommonRadioButton.Checked = $True ; $CommonPathTextField.Text = $pNewPath ; 
        $CommonPathTextField.BackColor = $BackgroundColor ; $IndividualPathTextField.BackColor = ""
        $Script:v_Root_CommonRepoFolder = $pNewPath
        $find = "^[\$]v_Root_CommonRepoFolder"
        $replace = "`$v_Root_CommonRepoFolder = ""$pNewPath"""
    } else {
        $IndividualRadioButton.Checked = $True ; $IndividualPathTextField.Text = $pNewPath
        $IndividualPathTextField.BackColor = $BackgroundColor ; $CommonPathTextField.BackColor = ""
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
        $true { $ca_FlagFilePath = $pPath+'\'+$MyAddSoftwareFolder+'\'+$pSysID } # $true
        $false { $ca_FlagFilePath = $pPath+'\'+$pModelName+'\'+$MyAddSoftwareFolder+'\'+$pSysID } # $false
    } # switch ( $pCommonFlag )

    # create the file and add selected Softpaqs content

    if ( -not (Test-Path $ca_FlagFilePath) ) { 
        CMTraceLog -message "Creating AddOn Flag File: $ca_FlagFilePath"
        New-Item $ca_FlagFilePath -Force > $null 
    } 
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
        $true { $ma_FlagFilePath = $pPath+'\'+$MyAddSoftwareFolder+'\'+$pSysID }
        $false { $ma_FlagFilePath = $pPath+'\'+$pModelName+'\'+$MyAddSoftwareFolder+'\'+$pSysID }
    } # switch ( $pCommonFlag )
    $ma_FlagFilePathBK = $ma_FlagFilePath+'_BK'
    #$ma_RepoPathSysID = $pPath+'\'+$pSysID+'_'+$pModelName+'\'+$MyAddSoftwareFolder+'\'+$pSysID
    #$ma_RepoPathSysIDBK = $ma_RepoPathSysID+'_BK'

    if ( $pUseFile ) {
        if ( -not (Test-Path $ma_FlagFilePath) ) {
            if ( Test-Path $ma_FlagFilePathBK ) {
                Move-Item $ma_FlagFilePathBK $ma_FlagFilePath -Force
            } else {
                if ( $v_DebugMode ) { CMTraceLog -Message "... calling Create_AddOnFlagFile()" -Type $TypeNorm }
                Create_AddOnFlagFile $pPath $pSysID $pModelName $pAddOns $pCommonFlag
            } # else if ( Test-Path $ma_FlagFilePathBK )
        } # if ( -not (Test-Path $ma_FlagFilePath) )
        if ( $v_DebugMode ) { CMTraceLog -Message "... calling Edit_AddonFlagFile() - for additional Softpaqs" -Type $TypeNorm }
        Edit_AddonFlagFile $ma_FlagFilePath $pSysID
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
	param( $pFlagFilePath, $pPlatformID )

    #$pSysID = Split-Path $pFlagFilePath -leaf
    $ea_FileContents = Get-Content $pFlagFilePath

    $ea_EntryFormWidth = 400 ; $ea_EntryFormHeigth = 400 ; $ea_FieldOffset = 20 ; $ea_FieldHeight = 20 ; $ea_PathFieldLength = 200

    if ( $v_DebugMode ) { Write-Host 'Manage AddOn Entries Form' }
    $ea_SoftpaqsForm = New-Object System.Windows.Forms.Form
    $ea_SoftpaqsForm.MaximizeBox = $False ; $ea_SoftpaqsForm.MinimizeBox = $False #; $EntryForm.ControlBox = $False
    $ea_SoftpaqsForm.Text = "Additional Softpaqs (Platform $($pPlatformID))"
    $ea_SoftpaqsForm.Width = $ea_EntryFormWidth ; $ea_SoftpaqsForm.height = $ea_EntryFormHeigth ; $ea_SoftpaqsForm.Autosize = $true
    $ea_SoftpaqsForm.StartPosition = 'CenterScreen' ; $ea_SoftpaqsForm.Topmost = $true

    # -----------------------------------------------------------
    $ea_EntryIDLabel = New-Object System.Windows.Forms.Label
    $ea_EntryIDLabel.Text = "Name"
    $ea_EntryIDLabel.location = New-Object System.Drawing.Point($ea_FieldOffset,$ea_FieldOffset)        # (from left, from top)
    $ea_EntryIDLabel.Size = New-Object System.Drawing.Size(60,20)                                       # (width, height)

    $ea_SoftpaqEntry = New-Object System.Windows.Forms.TextBox
    $ea_SoftpaqEntry.Text = "" ; $ea_SoftpaqEntry.Name = "Softpaq Name"
    $ea_SoftpaqEntry.Multiline = $false ; $ea_SoftpaqEntry.ReadOnly = $False
    $ea_SoftpaqEntry.location = New-Object System.Drawing.Point(($ea_FieldOffset+70),($ea_FieldOffset-4)) # (from left, from top)
    $ea_SoftpaqEntry.Size = New-Object System.Drawing.Size($ea_PathFieldLength,$ea_FieldHeight)         # (width, height)    

    $ea_SoftpaqMsg = New-Object System.Windows.Forms.Label
    $ea_SoftpaqMsg.Text = "Add Softpaq to maintain in the repository" ; $ea_SoftpaqEntry.Name = "Softpaq Message"
    $ea_SoftpaqMsg.ForeColor = 'blue'
    $ea_SoftpaqMsg.location = New-Object System.Drawing.Point($ea_FieldOffset,($ea_EntryFormHeigth-198))   # (from left, from top)    
    $ea_SoftpaqMsg.Size = New-Object System.Drawing.Size(($ea_PathFieldLength+60), ($ea_FieldHeight+200))           # (width, height)    

    $ea_AddButton = New-Object System.Windows.Forms.Button
    $ea_AddButton.Location = New-Object System.Drawing.Point(($ea_PathFieldLength+$ea_FieldOffset+80),($ea_FieldOffset-6))
    $ea_AddButton.Size = New-Object System.Drawing.Size(75,23)
    $ea_AddButton.Text = 'Add'

    $ea_SoftpaqListRetrieved = $False
    $ea_AddButton_AddClick = {
        if ( $ea_SoftpaqEntry.Text ) {
            # check for Softpaq names that match the pattern of the typed entry
            if ( -not $ea_SoftpaqListRetrieved ) {
                # retrieved the Softpaq list only once to avoid multiple calls to the CMSL command
                $ea_Softpaqs = Get-SoftpaqList -Platform $pPlatformID -os $Script:v_OS -OsVer $Script:v_OSVER -ErrorAction SilentlyContinue
                $ea_SoftpaqListRetrieved = $true
            } # if ( -not $ea_SoftpaqListRetrieved )

            # find all Softpaqs that match the entry text from the Softpaq list retrieved - could be more than one
            $ea_SoftpaqList = $ea_Softpaqs | Where-Object { $_.Name -like "*$($ea_SoftpaqEntry.Text)*" } | Sort-Object Name

            if ( $ea_SoftpaqList.Count -eq 0 ) {
                $ea_SoftpaqMsg.Text = "No matching Softpaqs found for '$($ea_SoftpaqEntry.Text)'"
            } else {
                # make sure the new Softpaq entry is not already in the flag file contents
                $ea_entryMatched = $ea_FileContents | ForEach-Object { $_.Trim() } | Where-Object { $_ -match $ea_SoftpaqEntry.Text } # trim empty lines
                if ( $ea_entryMatched ) {
                    $ea_SoftpaqMsg.Text = "Item '$($ea_SoftpaqEntry.Text)' already exists in the list."
                } else {
                    $ea_SoftpaqList | foreach {
                        # add each matching Softpaq to the  list
                        $ea_EntryList.items.add($_.Name)
                        CMTraceLog -Message "... Adding Softpaq: '$($_.Name)'" -Type $TypeNorm
                     } # $ea_SoftpaqList | foreach ( $ea_SoftpaqList )
                } # else if ( $ea_SoftpaqEntry.Text -in $ea_FileContents )
                $ea_entries = @() ; foreach ( $iEntry in $ea_EntryList.items ) { $ea_entries += $iEntry }
                Set-Content $pFlagFilePath -Value $ea_entries      # reset the file with the list of AddOns entries
            } # else if ( $ea_SoftpaqList.Count -eq 0 )
            $ea_SoftpaqEntry.Clear()  # clear the textbox for next entry
            $ea_SoftpaqEntry.Focus()  # set focus back to the textbox   
        } else {
            $ea_SoftpaqMsg.Text = '... nothing to add' 
        } # else if ( $EntryModel.Text )
    } # $ea_AddButton_AddClick =
    $ea_AddButton.Add_Click( $ea_AddButton_AddClick )

    # -----------------------------------------------------------

    $ea_EntryList = New-Object System.Windows.Forms.ListBox
    $ea_EntryList.Name = 'Entries'
    $ea_EntryList.Autosize = $false
    $ea_EntryList.location = New-Object System.Drawing.Point($ea_FieldOffset,60)  # (from left, from top)
    $ea_EntryList.Size = New-Object System.Drawing.Size(($ea_EntryFormWidth-60),($ea_EntryFormHeigth/2-60)) # (width, height)
    # populate the ListBox with existing entries from the flag file
    foreach ( $iSoftpaqName in $ea_FileContents ) { $ea_EntryList.items.add($iSoftpaqName) }

    $ea_EntryList.Add_Click({
        #$AddSoftpaqList.items.Clear()
        #foreach ( $iName in $pSoftpaqList ) { $AddSoftpaqList.items.add($iName) }
    })
    #$ea_EntryList.add_DoubleClick({ return $ea_EntryList.SelectedItem })
    
    # -----------------------------------------------------------
    $ea_removeButton = New-Object System.Windows.Forms.Button
    $ea_removeButton.Location = New-Object System.Drawing.Point(($ea_EntryFormWidth-120),($ea_EntryFormHeigth-200))
    $ea_removeButton.Size = New-Object System.Drawing.Size(75,23)
    $ea_removeButton.Text = 'Remove'
    $ea_removeButton.add_Click({
        if ( $ea_EntryList.SelectedItem ) {
            $ea_EntryList.items.remove($ea_EntryList.SelectedItem)
            $ea_entries = @()
            foreach ( $iEntry in $ea_EntryList.items ) { $ea_entries += $iEntry }
            Set-Content $pFlagFilePath -Value $ea_entries      # reset the file with needed AddOns
        }
    }) # $ea_removeButton.add_Click
    $ea_okButton = New-Object System.Windows.Forms.Button
    $ea_okButton.Location = New-Object System.Drawing.Point(($ea_EntryFormWidth-120),($ea_EntryFormHeigth-80))
    $ea_okButton.Size = New-Object System.Drawing.Size(75,23)
    $ea_okButton.Text = 'OK'
    $ea_okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $ea_SoftpaqsForm.AcceptButton = $ea_okButton
    $ea_SoftpaqsForm.Controls.AddRange(@($ea_EntryIDLabel,$ea_SoftpaqEntry,$ea_AddButton,$ea_EntryList,$SpqLabel,$AddSoftpaqList, $ea_removeButton, $ea_SoftpaqMsg, $ea_okButton))
    # -------------------------------------------------------------------
    # Ask the user what model to add to the list
    # -------------------------------------------------------------------
    $ea_SoftpaqsForm.ShowDialog()

} # Function Edit_AddonFlagFile

#=====================================================================================
<#
    Function Remove_SinglePlatform
    Here we remove the platform from the repository by removing the flag file and the grid platform entry
    parameters
        $pGrid
        pGridIndex              # the row index of the platform to remove
        $pRepositoryFolder      # the repository folder to find the associated flag file (used maintain addons list)
        $pCommonFlag            # is this a common repository? (flag files are in the root of the repository)
#>
Function Remove_SinglePlatform {
    [CmdletBinding()]
        param( $pGrid, 
                $pGridIndex,
                $pRepositoryFolder, 
                $pCommonFlag )

    $rs_SysID = $pGrid.Rows[$pGridIndex].Cells[1].Value     # obtain the platform SysID from the grid row (2nd col)

    switch ( $pCommonFlag ) {
        $true {
            # we are dealing with a common repository, so we just remove the flag file
            # later when the repo is sync'e'd, the platform will be removed from the repo - since the row will be deleted in this function, we don't need to worry about the grid here
            # ... the flag file name has 4 characters (the motherboard/SysID)
            # the selected platform for removal from the repository will be removed from the grid prior to exit from this function
            $rs_FlagFilePath = $pRepositoryFolder+'\'+$MyAddSoftwareFolder+'\'+$rs_SysID  
            if ( $v_DebugMode ) { CMTraceLog -Message '-- common repo - searching for flag file: '+$rs_FlagFilePath -Type $TypeNorm }
            if ( Test-Path $rs_FlagFilePath ) {
                if ( $v_DebugMode ) { CMTraceLog -Message '     found SysID flag file: '+$rs_FlagFilePath+' - removing' -Type $TypeNorm }  # remove the flag file only
                Remove-Item -Path $rs_FlagFilePath      # -confirm
            } # if ( Test-Path $rs_FlagFilePath )
        } # $true
        $false {
            # here we are dealing with and individual repository,  possibly a repo with multiple SysIDs (repository folders are named by the model name)
            #  ... so we need to check if there are multiple SysIDs in the repository folder
            #  ... if there are multiple SysIDs, we just remove the flag file
            #  ... if there is only 1 SysID, we remove the flag file and rename the repository folder
            #  ... later when the repo is sync'e'd, the platform will be removed from the repo - since the row will be deleted in this function, we don't need to worry about the grid here
            $rs_SystemName = $pGrid.Rows[$pGridIndex].Cells[2].Value                    # obtain the platform name from the grid row (3rd col)
            $rs_RepoPath = $pRepositoryFolder+'\'+$rs_SystemName                        # the individual repository folder for the platform/model
            
            $rs_FlagFilePath = $rs_RepoPath+'\'+$MyAddSoftwareFolder+'\'+$rs_SysID      # the flag file path for the platform to be removed

            if ( Test-Path $rs_RepoPath ) {
                # flag file name has 4 characters (the motherboard/SysID) - find all of them, in case there are multiple SysIDs in the grid for the platform/model
                [array]$rs_FlagFiles = Get-ChildItem $rs_RepoPath'\'$MyAddSoftwareFolder | Where-Object {$_.BaseName.Length -eq 4} 

                if ( $rs_FlagFiles.count -gt 1  ) {
                    Rename-Item -Path $rs_RepoPath -NewName $rs_SystemName'_BK' -confirm          # just rename the repository file
                } else {
                    if ( $rs_FlagFiles.count -eq 1 -and ($rs_FlagFiles[0] -like $rs_SysID) ) {
                        Remove-Item -Path $rs_FlagFilePath -confirm                             # remove the flag file only
                    }
                } # else if ( $rs_FlagFiles.count -gt 1  ) {
            } # if ( Test-Path $rs_RepoPath )
        } # $false
    } # switch ( $pCommonFlag )

    # finally, remove the grid entry for this platform/model so that on next sync, the platform filters will be removed from the repository
    # ... this is done by the caller of this function, so we don't need to do it here
    $pGrid.Rows.RemoveAt($pGridIndex)

} # Function Remove_SinglePlatform

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
    # confirm there are selected entries/models to remove from the list, where left row cell is checked
    for ( $irow = 0 ; $irow -lt $pGrid.RowCount ; $irow++ ) {
        if ( $pGrid.Rows[$irow].Cells[0].Value ) {  # Add to remove list if 1st column is checked
            $rp_removeList += @{ SysID = $pGrid.Rows[$irow].Cells[1].Value ; SysName = $pGrid.Rows[$irow].Cells[2].Value }
        } # if ( $pGrid.Rows[$irow].Cells[0].Value )
    } # for ( $row = 0 ; $row -lt $pGrid.RowCount ; $row++ )

    if ( $rp_removeList ) {
        if ( $rp_removeList.count -eq 1 ) {
            $rp_Msg = "Remove the selected entry from the list so it will be deleted from the repository on next sync?"
        } else {
            $rp_Msg = "Remove the selected entries from the list so they will be deleted from the repository on next sync?"
        }
        if ( (AskUser $rp_Msg 1) -eq 'Ok') {    # 1 = OKCancel
            if ( $v_DebugMode ) { CMTraceLog -Message '... REMOVING from repository' -Type $TypeWarn }
            for ( $irow = $pGrid.RowCount-1 ; $irow -ge 0 ; $irow-- ) {
                if ( $pGrid.Rows[$irow].Cells[0].Value ) {
                    Remove_SinglePlatform $pGrid $irow $pRepositoryFolder $pCommonFlag
                    #$pGrid.Rows.RemoveAt($irow) 
                } # if ( $pGrid.Rows[$irow].Cells[0].Value )
            } # for ( $irow = 0 ; $irow -lt $pGrid.RowCount ; $irow++ )            
            if ( $v_DebugMode ) { CMTraceLog -Message '... calling Update_INIModelsFromGrid()' -Type $TypeWarn }
            Update_INIModelsFromGrid $pGrid $pRepositoryFolder            
        } else {
            $rp_removeList = $null
        } # if ( (AskUser $rp_Msg 1) -eq 'Ok')
    } else {
        CMTraceLog -Message  "There are no selected entries to remove" -Type $TypeNorm
    } # else if ( $rp_removeList )
    
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

    Set-Variable -Name ap_FormWidth -Option ReadOnly -Value 400 -Force
    Set-Variable -Name ap_FormHeigth -Option ReadOnly -Value 400 -Force
    Set-Variable -Name ap_Offset -Option ReadOnly -Value 20 -Force
    Set-Variable -Name ap_FieldHeight -Option ReadOnly -Value 20 -Force
    Set-Variable -Name ap_PathFieldLength -Option ReadOnly -Value 200 -Force

    $ap_EntryForm = New-Object System.Windows.Forms.Form
    $ap_EntryForm.MaximizeBox = $False ; $ap_EntryForm.MinimizeBox = $False #; $ap_EntryForm.ControlBox = $False
    $ap_EntryForm.Text = "Select a Device to Add"
    $ap_EntryForm.Width = $ap_FormWidth ; $ap_EntryForm.height = 400 ; $ap_EntryForm.Autosize = $true
    $ap_EntryForm.StartPosition = 'CenterScreen' ; $ap_EntryForm.Topmost = $true

    # ------------------------------------------------------------------------
    # find and add model entry
    # ------------------------------------------------------------------------
    $ap_EntryLabel = New-Object System.Windows.Forms.Label
    $ap_EntryLabel.Text = "Name"
    $ap_EntryLabel.location = New-Object System.Drawing.Point($ap_Offset,$ap_Offset) # (from left, from top)
    $ap_EntryLabel.Size = New-Object System.Drawing.Size(60,20)                   # (width, height)

    $ap_EntryModel = New-Object System.Windows.Forms.TextBox
    $ap_EntryModel.Text = ""       # start w/INI setting
    $ap_EntryModel.Multiline = $false 
    $ap_EntryModel.location = New-Object System.Drawing.Point(($ap_Offset+70),($ap_Offset-4)) # (from left, from top)
    $ap_EntryModel.Size = New-Object System.Drawing.Size($ap_PathFieldLength,$ap_FieldHeight)# (width, height)
    $ap_EntryModel.ReadOnly = $False
    $ap_EntryModel.Name = "Model Name"
    $ap_EntryModel.add_MouseHover($ShowHelp)
    $ap_SearchButton = New-Object System.Windows.Forms.Button
    $ap_SearchButton.Location = New-Object System.Drawing.Point(($ap_PathFieldLength+$ap_Offset+80),($ap_Offset-6))
    $ap_SearchButton.Size = New-Object System.Drawing.Size(75,23)
    $ap_SearchButton.Text = 'Search'
    $ap_SearchButton_AddClick = {
        if ( $ap_EntryModel.Text ) {
            $ap_AddEntryList.Items.Clear()
            $ap_Models = Get-HPDeviceDetails -Like -Name $ap_EntryModel.Text    # find all models matching entered text
            foreach ( $iModel in $ap_Models ) { 
                [void]$ap_AddEntryList.Items.Add($iModel.SystemID+'_'+$iModel.Name) 
            }
        } # if ( $ap_EntryModel.Text )
    } # $ap_SearchButton_AddClick =
    $ap_SearchButton.Add_Click( $ap_SearchButton_AddClick )

    $ap_ListBoxHeight = $ap_FormHeigth/2-60
    $ap_AddEntryList = New-Object System.Windows.Forms.ListBox
    $ap_AddEntryList.Name = 'Entries'
    $ap_AddEntryList.Autosize = $false
    $ap_AddEntryList.location = New-Object System.Drawing.Point($ap_Offset,60)  # (from left, from top)
    $ap_AddEntryList.Size = New-Object System.Drawing.Size(($ap_FormWidth-60),$ap_ListBoxHeight) # (width, height)
    $ap_AddEntryList.Add_Click({
        $ap_AddSoftpaqList.items.Clear()
        foreach ( $iName in $pSoftpaqList ) { $ap_AddSoftpaqList.items.add($iName) }
    })
    #$ap_AddEntryList.add_DoubleClick({ return $ap_AddEntryList.SelectedItem })
    
    # ------------------------------------------------------------------------
    # find and add initial softpaqs to the selected model
    # ------------------------------------------------------------------------
    $ap_SpqLabel = New-Object System.Windows.Forms.Label
    $ap_SpqLabel.Text = "Select initial Addon Softpaqs" 
    $ap_SpqLabel.location = New-Object System.Drawing.Point($ap_Offset,($ap_ListBoxHeight+60)) # (from left, from top)
    $ap_SpqLabel.Size = New-Object System.Drawing.Size(70,60)                   # (width, height)

    $ap_AddSoftpaqList = New-Object System.Windows.Forms.ListBox
    $ap_AddSoftpaqList.Name = 'Softpaqs'
    $ap_AddSoftpaqList.Autosize = $false
    $ap_AddSoftpaqList.SelectionMode = 'MultiExtended'
    $ap_AddSoftpaqList.location = New-Object System.Drawing.Point(($ap_Offset+70),($ap_ListBoxHeight+60))  # (from left, from top)
    $ap_AddSoftpaqList.Size = New-Object System.Drawing.Size(($ap_FormWidth-130),($ap_ListBoxHeight-40)) # (width, height)

    # ------------------------------------------------------------------------
    # show the dialog, and once user preses OK, add the model and create the flag file for addons
    # ------------------------------------------------------------------------
    $ap_okButton = New-Object System.Windows.Forms.Button
    $ap_okButton.Location = New-Object System.Drawing.Point(($ap_FormWidth-120),($ap_FormHeigth-80))
    $ap_okButton.Size = New-Object System.Drawing.Size(75,23)
    $ap_okButton.Text = 'OK' ; $ap_okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $ap_cancelButton = New-Object System.Windows.Forms.Button
    $ap_cancelButton.Location = New-Object System.Drawing.Point(($ap_FormWidth-200),($ap_FormHeigth-80))
    $ap_cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $ap_cancelButton.Text = 'Cancel' ; $ap_cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::CANCEL

    $ap_EntryForm.AcceptButton = $ap_okButton ; $ap_EntryForm.CancelButton = $ap_cancelButton
    $ap_EntryForm.Controls.AddRange(@($ap_EntryLabel,$ap_EntryModel,$ap_SearchButton,$ap_AddEntryList,$ap_SpqLabel,$ap_AddSoftpaqList, $ap_cancelButton, $ap_okButton))

    $ap_Result = $ap_EntryForm.ShowDialog()

    if ($ap_Result -eq [System.Windows.Forms.DialogResult]::OK) {        
        
        $ap_SelectedSysID = $ap_AddEntryList.SelectedItem.substring(0,4)
        $ap_SelectedModel = $ap_AddEntryList.SelectedItem.substring(5)  # name is after 'SysID_'
        
        [array]$ap_SelectedEntry = Get-HPDeviceDetails -Like -Name $ap_AddEntryList.SelectedItem

        #  check if the model is already in the grid, if so, don't add it again
        for ( $iRow = 0; $iRow -lt $pGrid.RowCount; $iRow++ ) {
            $ap_CurrEntrySysID = $pGrid.Rows[$iRow].Cells[1].value
            if ( $ap_CurrEntrySysID -like $ap_SelectedSysID ) {
                # the model is already in the grid, inform the user and return
                AskUser "This model is already in the Grid" 0 | Out-Null # 0="OK" button only
                $ap_platformExists = $True
            } # if ( $ap_CurrEntrySysID -like $ap_SelectedSysID )
        } # for ( $iRow = 0; $iRow -lt $pGrid.RowCount; $iRow++ )

        if ( $ap_platformExists ) { 
            $ap_SelectedEntry = $null 
        } else {           
            # add model to UI grid and initialize it as an HPIA repository
            [void]$pGrid.Rows.Add( @( $False, $ap_SelectedSysID, $ap_SelectedModel) )
            # get a list of selected additional softpaqs to add
            if ( $v_DebugMode ) { CMTraceLog -Message '... > Create_AddOnFlagFile()' -Type $TypeNorm }
            $ap_numEntries = Create_AddOnFlagFile $pRepositoryFolder $ap_SelectedSysID $ap_SelectedModel $ap_AddSoftpaqList.SelectedItems $pCommonFlag
            # check the AddOns cell if entries in flag file
            if ( $ap_numEntries -gt 0 ) {
                $pGrid.Rows[($pGrid.RowCount-1)].Cells['AddOns'].Value = $True
            }
            if ( $v_DebugMode ) { CMTraceLog -Message '... > Update_INIModelsFromGrid()' -Type $TypeNorm }
            Update_INIModelsFromGrid $pGrid $pRepositoryFolder
        } # else if ( $ap_platformExists )

    } # if ($ap_Result -eq [System.Windows.Forms.DialogResult]::OK)

    return $ap_SelectedEntry

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
    $Script:au_return = $null
    $au_FormWidth = 300 ; $au_FormHeight = 200

    $au_form = New-Object System.Windows.Forms.Form
    $au_form.Text = ""
    $au_form.Width = $au_FormWidth ; $au_form.height = $au_FormHeight ; $au_form.StartPosition = 'CenterScreen'
    $au_form.Topmost = $true

    $au_MsgBox = New-Object System.Windows.Forms.TextBox    
    $au_MsgBox.Name = "Ask User" ; $au_MsgBox.Text = $pText       # start w/INI setting
    $au_MsgBox.BorderStyle = 0 ; $au_MsgBox.Multiline = $true ; $au_MsgBox.TabStop = $false ; $au_MsgBox.ReadOnly = $true
    $au_MsgBox.location = New-Object System.Drawing.Point(20,20) # (from left, from top)
    $au_MsgBox.Size = New-Object System.Drawing.Size(($au_FormWidth-60),($au_FormHeight-120))# (width, height)
    #$au_MsgBox.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",($TextBox.Font.Size+4),[System.Drawing.FontStyle]::Regular)
    $au_MsgBox.Font = New-Object System.Drawing.Font("Verdana",($TextBox.Font.Size+2),[System.Drawing.FontStyle]::Regular)    

    #----------------------------------------------------------------------------------
    # Create Buttons at the bottom of the dialog
    #----------------------------------------------------------------------------------
    $au_button1 = New-Object System.Windows.Forms.Button
    $au_button1.Text = $au_askbutton1
    $au_button1.Location = New-Object System.Drawing.Point(($au_FormWidth-200),($au_FormHeight-80))    # (from left, from top)
    $au_button1.add_click( { $Script:au_return = $au_askbutton1 ; $au_form.Close() ; $au_form.Dispose() } )

    $au_button2 = New-Object System.Windows.Forms.Button
    $au_button2.Text = $au_askbutton2
    $au_button2.Location = New-Object System.Drawing.Point(($au_FormWidth-120),($au_FormHeight-80))    # (from left, from top)
    $au_button2.add_click( { $Script:au_return = $au_askbutton2 ; $au_form.Close() ; $au_form.Dispose() } )
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
            "OS_Version"        {$tip = "Select the Windows version to maintain in the repository"}
            "OS"                {$tip = "Select the Windows OS to maintain in the repository"}
            "Keep Filters"      {$tip = "Do NOT erase previous product selection filters"}
            "Continue on 404"   {$tip = "Continue Sync evern with Error 404, missing files"}
            "Individual Paths"  {$tip = "Path to Head of Individual platform repositories"}
            "Common Path"       {$tip = "Path to Common/Shared platform repository"}
            "Models Table"      {$tip = "HP Models table to Sync repository(ies) to"}
            "Check All"         {$tip = "This check selects all Platforms and categories"}
            "Sync"              {$tip = "Syncronize repository for selected items from HP cloud"}
            'Use INI List'      {$tip = 'Reset the Grid from the INI file $HPModelsTable list'}
            'Show Filters'      {$tip = 'Show list of all current Repository filters'}
            'Add Model'         {$tip = 'Find and add a model to the current list in the Grid'}
        } # Switch ($this.name)
        $CMForm_tooltip.SetToolTip($this,$tip)
    } #end ShowHelp

    if ( $v_DebugMode ) { Write-Host 'creating Form' }
    $CM_form = New-Object System.Windows.Forms.Form
    $CM_form.Text = "HPIARepo_Downloader v$($ScriptVersion)"
    $CM_form.Width = $FormWidth ; $CM_form.height = $FormHeight ; $CM_form.Autosize = $true
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
    #$OSHeader = 'Imported Repository'

    if ( $v_DebugMode ) { Write-Host 'creating Shared radio button' }
    $IndividualRadioButton = New-Object System.Windows.Forms.RadioButton
    $IndividualRadioButton.Location = '10,14'
    $IndividualRadioButton.Add_Click( {
            $Script:v_CommonRepo = $False
            $find = "^[\$]v_CommonRepo"
            $replace = "`$v_CommonRepo = `$$Script:v_CommonRepo" 
            Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - INI setting set: ''$($replace)''"
            Import_IndividualRepos $dataGridView $IndividualPathTextField.Text            
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
        Update_GridFromINI $dataGridView $Script:HPModelsTable
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
    $AddModelButton.Text = "Add`nModel"
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
    if ( $v_DebugMode ) { Write-Host 'creating a Remove Model button' }
    $RemoveModelButton = New-Object System.Windows.Forms.Button
    $RemoveModelButton.Width = 80 ; $RemoveModelButton.Height = 35
    $RemoveModelButton.Location = New-Object System.Drawing.Point(($LeftOffset+270),($FormHeight-427))    # (from left, from top)
    $RemoveModelButton.Text = 'Remove Model'
    $RemoveModelButton.Name = 'Remove Model'
    $RemoveModelButton.add_Click( { 
        if ( $Script:v_CommonRepo ) { $mf_RepoPath = $CommonPathTextField.Text } else { $mf_RepoPath = $IndividualPathTextField.Text }
        (Remove_Platforms $DataGridView $mf_RepoPath $Script:v_CommonRepo) | ForEach-Object {
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
    $buttonSync.Text = 'SYNC'
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
                if ( $Script:s_ModelAdded ) { Update_INIModelsFromGrid $dataGridView $CommonPathTextField }   # $True = head of common repository
            } else {
                sync_individualRepositories $dataGridView $IndividualPathTextField.Text $lCheckedListArray $Script:s_ModelAdded
                if ( $Script:s_ModelAdded ) { Update_INIModelsFromGrid $dataGridView $IndividualPathTextField.Text } 
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

    #$TextBoxFontDefault =  $TextBox.Font    # save the default font
    #$TextBoxFontDefaultSize = $TextBox.Font.Size

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
            $find = "^[\$]v_DebugMode"
            $replace = "`$v_DebugMode = `$$Script:v_DebugMode"                   # set up the replacing string to either $false or $true from ini file
            Mod_INISetting $IniFIleFullPath $find $replace "$($IniFile) - changing setting: ''$($replace)''"
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
    (Get-Content $IniFIleFullPath) | 
        Foreach-Object { 
            if ($_ -match $find) {               # we found the variable
                if ( $_ -match '\$true' ) { 
                    $m_CommonPath = $CommonPathTextField.Text
                    if ( ([string]::IsNullOrEmpty($m_CommonPath)) -or (-not (Test-Path $m_CommonPath)) ) {
                        Write-Host "Common Repository Path from INI file not found: $($m_CommonPath) - Will Create" -ForegroundColor Brown
                        Init_Repository $m_CommonPath $True                 # $True = make it a HPIA repo
                    }   
                    Import_Repository $dataGridView $m_CommonPath $Script:HPModelsTable
                    $CommonRadioButton.Checked = $true                     # set the visual default from the INI setting
                    $CommonPathTextField.BackColor = $BackgroundColor
                } else { 
                    if ( [string]::IsNullOrEmpty($IndividualPathTextField.Text) ) {
                        Write-Host "Individual Repository field is empty" -ForegroundColor Red
                    } else {
                        Import_IndividualRepos $dataGridView $IndividualPathTextField.Text
                        $IndividualRadioButton.Checked = $true 
                        $IndividualPathTextField.BackColor = $BackgroundColor
                    }
                } # else if ( $_ -match '\$true' )
                if ( $Script:v_KeepFilters ) {
                    AskUser "NOTE: `'Kepp Prev OS Filters`' is enabled. Existing repository filters will not be removed" 0 | out-null # 0 = Ok dialog
                }
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