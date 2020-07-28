<#
    HP Image Assistant and Softpaq Repository Downloader
    Version 1.00
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
        Code cleanup of Function sync_repositories 
        changed 'use single repository' checkbox to radio buttons on path fields
        Added 'Distribute SCCM Packages' ($DistributeCMPackages) variable use in INI file
            -- when selected, sends command to CM to Distribute packages
    Version 1.30
        Added ability to sync specific softpaqs by name - listed in INI file
            -- added SqName entry to $HPModelsTable list to hold special softpaqs needed/model
#>
param(
	[Parameter(Mandatory = $false,Position = 1,HelpMessage = "Application")]
	[ValidateNotNullOrEmpty()]
	[ValidateSet("sync")]
	$RunMethod = "sync"
)
$ScriptVersion = "1.30 (7/27/2020)"

# get the path to the running script, and populate name of INI configuration file
$scriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

#--------------------------------------------------------------------------------------
$IniFile = "HPIARepo_ini.ps1"                        # assume this INF file in same location as script
$IniFIleFullPath = "$($ScriptPath)\$($IniFile)"

. $IniFIleFullPath                                   # source the code in the INI file      

#--------------------------------------------------------------------------------------
#Script Vars Environment Specific loaded from INI.ps1 file

$CMConnected = $false                                # is a connection to SCCM established?
$SiteCode = $null

$AddSoftware = '.ADDSOFTWARE'                        # sub-folders where named Softpaqs will reside

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

    switch ( $pmsgType )
    {
       -1 { $pTextBox.SelectionColor = "Red" }                  # Error
        1 { $pTextBox.SelectionColor = "Black" }                # default color is black
        2 { $pTextBox.SelectionColor = "Brown" }                # Warning
        4 { $pTextBox.SelectionColor = "Orange" }               # Debug Output
        5 { $pTextBox.SelectionColor = "Green" }                # success details
        10 { $pTextBox.SelectionColor = "Black" }               # do NOT add \newline to message output
    } # switch ( $pmsgType )

    if ( $pmsgType -eq $TypeDebug ) {
        $pMessage = '{dbg}'+$pMessage
    }

    # message Tpye = 10/$TypeNeNewline prevents a nl so next output is written contiguous

    if ( $pmsgType -eq $TypeNoNewline ) {
        $pTextBox.AppendText("$($pMessage) ")
    } else {
        $pTextBox.AppendText("$($pMessage) `n")
    }
    $pTextBox.Refresh()
    $pTextBox.ScrollToCaret()

} # Function OutToForm

#=====================================================================================
<#
    Function Load_HPModule
        The function will test if the HP Client Management Script Library is loaded
        and attempt to load it, if possible
#>
function Load_HPModule {

    if ( $DebugMode ) { CMTraceLog -Message "> Load_HPModule" -Type $TypeNorm }
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
                Install-PackageProvider -Name NuGet -ForceBootstrap
                Install-Module -Name PowerShellGet -Force

                if ( $DebugMode ) { write-host "Installing and Importing Module $m." }
                CMTraceLog -Message "Installing and Importing Module $m." -Type $TypSuccess
                Install-Module -Name $m -Force -SkipPublisherCheck -AcceptLicense -Verbose -Scope CurrentUser
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

    if ( $DebugMode ) { CMTraceLog -Message "< Load_HPModule" -Type $TypeNorm }

} # function Load_HPModule

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
            #$RepoShareMain = "\\$($FileServerName)\share\softpaqs\HPIARepo"

            if (Test-Path $lCMInstall) {
        
                try { Test-Connection -ComputerName "$FileServerName" -Quiet

                    CMTraceLog -Message " ...Connected" -Type $TypeSuccess 
                    $boolConnectionRet = $True
                }
                catch {
	                CMTraceLog -Message "Not Connected to File Server, Exiting" -Type $TypeError 
                }
            } else {
                CMTraceLog -Message "CM Installation path NOT FOUND: '$lCMInstall'" -Type $TypeError 
            } # else
        }
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
    Function CM_RepoUpdate
#>
Function CM_RepoUpdate {
    [CmdletBinding()]
	param( $pModelName, $pModelProdId, $pRepoPath )                             

    $pCurrentLoc = Get-Location

    CMTraceLog -Message "> CM_RepoUpdate" -Type $TypeNorm
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

    Set-Location -Path $pCurrentLoc

    CMTraceLog -Message "< CM_RepoUpdate" -Type $TypeNorm

} # Function CM_RepoUpdate

#=====================================================================================
<#
    Function Get_LogEntries
        obtains log entries from the ..\.Repository\activity.log file from last sync
#>
Function Get_LogEntries {
    [CmdletBinding()]
	param( $pRepoFolder )

    $lDotRepository = "$($pRepoFolder)\.Repository"
    $lLastSync = $null
    $lCurrRepoLine = 0
    $lLastSyncLine = 0

    if ( Test-Path $lDotRepository ) {

        #--------------------------------------------------------------------------------
        # find the last Sync started entry line
        #--------------------------------------------------------------------------------
        $lRepoLogFile = "$($lDotRepository)\activity.log"

        if ( Test-Path $lRepoLogFile ) {
            #(Get-Content $lRepoLogFile) | Foreach-Object {if ($_ -match $find) {$replace} else {$_}} #| Set-Content $IniFIleFullPath
            $find = 'sync has started'

            (Get-Content $lRepoLogFile) | 
                Foreach-Object { 
                    $lCurrRepoLine++ 
                    if ($_ -match $find) { $lLastSync = $_ ; $lLastSyncLine = $lCurrRepoLine } 
                } # Foreach-Object
            CMTraceLog -Message "... [CMSL activity.log - last sync] $lLastSync (line $lLastSyncLine)" -Type $TypeWarn
        } # if ( Test-Path $lRepoLogFile )

        #--------------------------------------------------------------------------------
        # now, $lLastSyncLine holds the log's line # where the last sync started
        #--------------------------------------------------------------------------------
        if ( $lLastSync ) {

            $lLogFile = Get-Content $lRepoLogFile

            for ( $i = 0; $i -lt $lLogFile.Count; $i++ ) {
                if ( $i -ge ($lLastSyncLine-1) ) {
                    if ( ($lLogFile[$i] -match 'done downloading exe') -or (($lLogFile[$i] -match 'already exists')) ) {
                        CMTraceLog -Message "... [CMSL activity.log - Softpaq update] $($lLogFile[$i])" -Type $TypeWarn
                    }
                }
            } # for ( $i = 0; $i -lt $lLogFile.Count; $i++ )
        } # if ( $lLastSync )

    } # if ( Test-Path $lDotRepository )

} # Function Get_LogEntries

#=====================================================================================
<#
    Function init_repository
        This function will create a repository foldern and initialize it for HPIA, if necessary
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

    Set-Location $lCurrentLoc

    if ( $DebugMode ) { CMTraceLog -Message "< init_repository" -Type $TypeNorm }
    return $retRepoCreated

} # Function init_repository

#=====================================================================================

Function download_softpaqs_by_name {
    [CmdletBinding()]
	param( $pFolder,
            $pProdCode )

    if ( $DebugMode ) { CMTraceLog -Message  '> download_softpaqs_by_name' }

    $HPModelsTable  | 
        ForEach-Object {
            if ( $_.ProdCode -match $pProdCode ) {                              # let's match the Model to download for
                
                $lAddSoftpaqsFolder = "$($pFolder)\$($AddSoftware)"
                Set-Location $lAddSoftpaqsFolder
                CMTraceLog -Message "... Retrieving Named Softpaqs for ProdCode:$pProdCode" -Type $TypeNorm

                ForEach ( $Softpaq in $_.SqName ) {
                    
                    Get-SoftpaqList -platform $pProdCode -os $OS -osver $OSVER | Where-Object { $_.Name -eq $Softpaq } | 
                        ForEach-Object { 
                            if ($_.SSM -eq $true) {
                                CMTraceLog -Message "... [Get Softpaq] - $($_.Id) ''$Softpaq''" -Type $TypeWarn
                                $ret = (Get-Softpaq $_.Id) 6>&1
                                if ( $DebugMode ) { CMTraceLog -Message  "... Get-Softpaq: $($ret)"  -Type $TypeWarn }
                                $ret = (Get-SoftpaqMetadataFile $_.Id) 6>&1
                                if ( $DebugMode ) { CMTraceLog -Message  "... Get-SoftpaqMetadataFile done: $($ret)"  -Type $TypeWarn }
                            } # if ($_.SSM -eq $true)
                        } # ForEach-Object

                } # ForEach ( $Softpaq in $_.SqName )
<#
                # next, copy all softpaqs in $AddSoftware subfolder to the repository (since it got clearn up by CMSL's "Invoke-RepositoryCleanup"

                if ( Test-Path $lAddSoftpaqsFolder ) {
                    CMTraceLog -Message "... Copying Added/Needed softpaqs to Repository"
                    $lRobocopySource = $lAddSoftpaqsFolder
                    $lRobocopyDest = "$($pFolder)"
                    $lRobocopyArg = '"'+$lRobocopySource+'"'+' "'+$lRobocopyDest+'"'
                    $RobocopyCmd = "robocopy.exe"
                    Start-Process -FilePath $RobocopyCmd -ArgumentList $lRobocopyArg -Wait     
                    Write-Host 'robocopy arg='$lRobocopyArg
                }
#>
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

    $pCurrentLoc = Get-Location

    if ( Test-Path $pFolder ) {

        #--------------------------------------------------------------------------------
        # update repository softpaqs with sync command and then cleanup
        #--------------------------------------------------------------------------------
        Set-Location -Path $pFolder

        CMTraceLog -Message  '... invoking repository sync - please wait !!!' -Type $TypeNorm
        $lRes = (invoke-repositorysync 6>&1)

        # find what sync'd from the CMSL log file for this run
        Get_LogEntries $pFolder

        CMTraceLog -Message  '... invoking repository cleanup ' -Type $TypeNoNewline 
        $lRes = (invoke-RepositoryCleanup 6>&1)
        CMTraceLog -Message "... $($lRes)" -Type $TypeWarn

        # next, copy all softpaqs in $AddSoftware subfolder to the repository (since it got clearn up by CMSL's "Invoke-RepositoryCleanup"

        CMTraceLog -Message "... Copying named softpaqs to Repository"
        $lRobocopySource = "$($pFolder)\$($AddSoftware)"
        $lRobocopyDest = "$($pFolder)"
        $lRobocopyArg = '"'+$lRobocopySource+'"'+' "'+$lRobocopyDest+'"'
        $RobocopyCmd = "robocopy.exe"
        Start-Process -FilePath $RobocopyCmd -ArgumentList $lRobocopyArg -Wait     

    } # if ( Test-Path $pFolder )

    Set-Location $pCurrentLoc

} # Function Sync_and_Cleanup_Repository

#=====================================================================================
<#
    Function Update_Repository_and_Grid
        for the selected model, 
            - remove all filters
            - add filters as selected in UI
            - invoke sync and cleanup function on the folder 
#>
Function Update_Repository_and_Grid {
[CmdletBinding()]
	param( $pModelRepository,
            $pModelID,
            $pRow,
            $pmodelChecked )

    $pCurrentLoc = Get-Location

    # move to location of Repository to use CMSL repo commands
    set-location $pModelRepository

    #--------------------------------------------------------------------------------
    # clean up filters for the current model in this loop
    #--------------------------------------------------------------------------------

    if ( $debugMode ) { CMTraceLog -Message  "... removing filters in: $($pModelRepository)" -Type $TypeWarn }
    $lres = (Remove-RepositoryFilter -platform $pModelID -yes 6>&1)      
    if ( $debugMode ) { CMTraceLog -Message "... removed filter: $($lres)" -Type $TypeWarn }

    if ( $pmodelChecked ) {
        #--------------------------------------------------------------------------------
        # update filters - every category checked for the current model'
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
        # update repository path for this model in the grid (col 8 is links col)
        #--------------------------------------------------------------------------------
        $datagridview[8,$pRow].Value = $pModelRepository
    } else {
        $datagridview[8,$pRow].Value = ""
    } # else if ( $pmodelChecked ) {

    Set-Location -Path $pCurrentLoc

} # Update_Repository_and_Grid

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
            $pCheckedItemsList)                                      # array of rows selected
    
    $pCurrentLoc = Get-Location
    CMTraceLog -Message "> Sync_Common_Repository - Common Folder" -Type $TypeNorm
    if ( $DebugMode ) { CMTraceLog -Message "... for OS version: $($OSVER)" -Type $TypeDebug }

    $lMainRepo =  $RepoShareCommon                               # 
    init_repository $lMainRepo $true                             # make sure Main repo folder exists, or create it - no init
    CMTraceLog -Message  "... Common repository selected: $($lMainRepo)" -Type $TypeNorm

    #--------------------------------------------------------------------------------
    if ( $DebugMode ) { CMTraceLog -Message "... stepping through selected models" -Type $TypeDebug }

    # loop through every Model in the list

    for ( $i = 0; $i -lt $pModelsList.RowCount; $i++ ) {
        
        $lModelId = $pModelsList[1,$i].Value                            # column 1 has the Model/Prod ID
        $lModelName = $pModelsList[2,$i].Value                          # column 2 has the Model name

        # if model entry is checked, we need to create a repository, but ONLY if $CommonRepo = $false
        if ( $i -in $pCheckedItemsList ) {

            CMTraceLog -Message "... Updating model: $lModelName"

            Update_Repository_and_Grid $lMainRepo $lModelId $i $true    # $true means 'add filters'

            #--------------------------------------------------------------------------------
            # update SCCM Repository package, if user allows
            #--------------------------------------------------------------------------------
            if ( $Script:UpdateCMPackages ) {
                CM_RepoUpdate $lModelName $lModelId $lMainRepo
            }

            #--------------------------------------------------------------------------------
            # update specific Softpaqs as defined in INI file for the current model
            #--------------------------------------------------------------------------------
            download_softpaqs_by_name  $lMainRepo $lModelID

        } else {

            # make sure previous filters from 'unselected' models are removed from repository
            Update_Repository_and_Grid $lMainRepo $lModelId $i $false   # $false means 'do not add filters, just remove them'

        } # if ( $i -in $pCheckedItemsList )

    } # for ( $i = 0; $i -lt $pModelsList.RowCount; $i++ )

    # we are done checking every model for filters, so cleanup

    Sync_and_Cleanup_Repository $lMainRepo

    #--------------------------------------------------------------------------------
    Set-Location -Path $pCurrentLoc

    CMTraceLog -Message "< Sync_Common_Repository DONE" -Type $TypeSuccess

} # Function Sync_Common_Repository

#=====================================================================================
<#
    Function Sync_Repositories
        for every selected model, go through every repository by model
            - ensure there is a valid repository
            - remove all filters
            - add filters as selected in UI
            - invoke sync and cleanup function on the folder 
#>
Function Sync_Repositories {
[CmdletBinding()]
	param( $pModelsList,                                             # array of row lines that are checked
            $pCheckedItemsList)                                      # array of rows selected
    
    $pCurrentLoc = Get-Location
    CMTraceLog -Message "> Sync_Repositories - Individual Folders" -Type $TypeNorm
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

            CMTraceLog -Message "... Updating model: $lModelName"

            $lTempRepoFolder = "$($lMainRepo)\$($lModelName)"     # this is the repo folder for this model
            init_repository $lTempRepoFolder $true
            
            Update_Repository_and_Grid $lTempRepoFolder $lModelId $i $true    # $true means 'add filters'
            
            #--------------------------------------------------------------------------------
            # update specific Softpaqs as defined in INI file for the current model
            #--------------------------------------------------------------------------------
            download_softpaqs_by_name  $lTempRepoFolder $lModelID

            #--------------------------------------------------------------------------------
            # now sync up and cleanup this repository - if common repo, leave this for later
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

    #--------------------------------------------------------------------------------
    Set-Location -Path $pCurrentLoc

    CMTraceLog -Message "< Sync_Repositories DONE" -Type $TypeSuccess

} # Function Sync_Repositories

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
    Function get_filters
        Retrieves category filters from the share for each selected model
        ... and populates the Grid appropriately

#>
Function get_filters {
    [CmdletBinding()]
	param( $pDataGridList )                             # array of row lines that are checked

    $pCurrentLoc = Get-Location

    if ( $DebugMode ) { CMTraceLog -Message '> get_filters' -Type $TypeNorm }
    
    clear_datagrid $pDataGridList

    #--------------------------------------------------------------------------------
    # find out if the share exists, if not, just return
    #--------------------------------------------------------------------------------
    if ( (Test-Path $RepoShareMain) -or (Test-Path $RepoShareCommon) ) {
        CMTraceLog -Message '... Refreshing Grid ...' -Type $TypeNorm

        #--------------------------------------------------------------------------------
        if ( $Script:CommonRepo ) {
            # see if the repository was configured for HPIA

            if ( !(Test-Path "$($RepoShareCommon)\.Repository") ) {
                CMTraceLog -Message "... Repository Folder not initialized" -Type $TypeWarn
                CMTraceLog -Message '> get_filters - Done' -Type $TypeNorm
                return
            } 
            set-location $RepoShareCommon            

            $lProdFilters = (get-repositoryinfo).Filters

            foreach ( $filterSetting in $lProdFilters ) {

                if ( $DebugMode ) { CMTraceLog -Message "... populating filter categories ''$($lModelName)''" -Type $TypeDebug }

                # check each row SysID against the Filter Platform ID
                for ( $i = 0; $i -lt $pDataGridList.RowCount; $i++ ) {

                    if ( $filterSetting.platform -eq $pDataGridList[1,$i].value) {
                        # we matched the row/SysId with the Filter Platform IF, so let's add each category in the filter
                        foreach ( $cat in  ($filterSetting.category.split(' ')) ) {
                            $pDataGridList.Rows[$i].Cells[$cat].Value = $true
                        }
                        $pDataGridList[0,$i].Value = $true
                        $pDataGridList[8,$i].Value = $RepoShareCommon
                    } # if ( $filterSetting.platform -eq $pDataGridList[1,$i].value)

                } # for ( $i = 0; $i -lt $pDataGridList.RowCount; $i++ )
            } # foreach ( $platform in $lProdFilters )

        } else {

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

                        CMTraceLog -Message "... populating filter categories ''$($lModelName)''" -Type $TypeDebug
                        
                        foreach ( $cat in  ($lProdFilters.category.split(' ')) ) {
                            $pDataGridList.Rows[$i].Cells[$cat].Value = $true
                        }
                        #--------------------------------------------------------------------------------
                        # show repository path for this model (model is checked) (col 8 is links col)
                        #--------------------------------------------------------------------------------
                        $pDataGridList[0,$i].Value = $true
                        $pDataGridList[8,$i].Value = $lTempRepoFolder

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
    Function CreateForm
    This is the MAIN function with a Gui that sets things up for the user
#>
Function CreateForm {
    
    Add-Type -assembly System.Windows.Forms

    $LeftOffset = 20
    $TopOffset = 20
    $FieldHeight = 20
    $FormWidth = 800
    $FormHeight = 600
    if ( $DebugMode ) { Write-Host 'creating Form' }
    $CM_form = New-Object System.Windows.Forms.Form
    $CM_form.Text = "CM_HPIARepo_Downloader v$($ScriptVersion)"
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
            if ( $DebugMode ) { CMTraceLog -Message 'modifying INI file with selected OSVER' -Type $TypeNorm }
            $find = "^[\$]OSVER"
            $replace = "`$OSVER = ""$($OSVERComboBox.Text)"""  
            (Get-Content $IniFIleFullPath) | Foreach-Object {if ($_ -match $find) {$replace} else {$_}} | Set-Content $IniFIleFullPath
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
            } # for

            if ( $updateCMCheckbox.checked ) {
                if ( ($Script:CMConnected = Test_CMConnection) ) {
                    if ( $DebugMode ) { CMTraceLog -Message 'Script connected to CM' -Type $TypeDebug }
                }
            }
            if ( $Script:CommonRepo ) {
                Sync_Common_Repository $dataGridView $lCheckedListArray
            } else {
                Sync_Repositories $dataGridView $lCheckedListArray
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
    $OSTextLabel.Text = "Win 10 Version:"
    $OSTextLabel.location = New-Object System.Drawing.Point(($LeftOffset+70),($TopOffset+4))    # (from left, from top)
    $OSTextLabel.Size = New-Object System.Drawing.Size(90,25)                               # (width, height)
    #$OSTextField = New-Object System.Windows.Forms.TextBox
    $OSVERComboBox = New-Object System.Windows.Forms.ComboBox
    $OSVERComboBox.Size = New-Object System.Drawing.Size(60,$FieldHeight)                  # (width, height)
    $OSVERComboBox.Location  = New-Object System.Drawing.Point(($LeftOffset+160), ($TopOffset))
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
    # Create 'Debug Mode' - checkmark
    #----------------------------------------------------------------------------------
    $DebugCheckBox = New-Object System.Windows.Forms.CheckBox
    $DebugCheckBox.Text = 'Debug Mode'
    $DebugCheckBox.UseVisualStyleBackColor = $True
    $DebugCheckBox.location = New-Object System.Drawing.Point(($LeftOffset+240),$TopOffset)   # (from left, from top)
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
    # add share info field
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
    $SharePathTextField.Size = New-Object System.Drawing.Size(290,$FieldHeight)             # (width, height)
    $SharePathTextField.ReadOnly = $true
    $SharePathTextField.Name = "Share_Path"

    $SharePathLabelSingle = New-Object System.Windows.Forms.Label
    $SharePathLabelSingle.Text = "Common"
    $SharePathLabelSingle.location = New-Object System.Drawing.Point(($LeftOffset+15),($TopOffset+20)) # (from left, from top)
    $SharePathLabelSingle.Size = New-Object System.Drawing.Size(55,20)                            # (width, height)
    #$SharePathLabelSingle.TextAlign = "Left"
    $SharePathSingleTextField = New-Object System.Windows.Forms.TextBox
    $SharePathSingleTextField.Text = "$RepoShareCommon"
    $SharePathSingleTextField.Multiline = $false 
    $SharePathSingleTextField.location = New-Object System.Drawing.Point(($LeftOffset+80),($TopOffset+15)) # (from left, from top)
    $SharePathSingleTextField.Size = New-Object System.Drawing.Size(290,$FieldHeight)             # (width, height)
    $SharePathSingleTextField.ReadOnly = $true
    $SharePathSingleTextField.Name = "Single_Share_Path"
    
    $ShareButton = New-Object System.Windows.Forms.RadioButton
    $ShareButton.Location = '10,14'

    $CommonButton = New-Object System.Windows.Forms.RadioButton
    $CommonButton.Location = '10,34'

    $PathsGroupBox = New-Object System.Windows.Forms.GroupBox
    $PathsGroupBox.location = New-Object System.Drawing.Point(($LeftOffset+340),($TopOffset))     # (from left, from top)
    $PathsGroupBox.Size = New-Object System.Drawing.Size(400,65)                              # (width, height)
    $PathsGroupBox.text = "HPIA Repository Paths - from $($IniFile):"

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
        CMTraceLog -Message "updating $($IniFile) setting: ''$($replace)''" -Type $TypeNorm
        (Get-Content $IniFIleFullPath) | Foreach-Object {if ($_ -match $find) {$replace} else {$_}} | Set-Content $IniFIleFullPath

        get_filters $dataGridView                                          # re-populate categories for available systems
        
    } # $CommonRepoRadio_Click

    $ShareButton.Add_Click( $CommonRepoRadio_Click )
    $CommonButton.Add_Click( $CommonRepoRadio_Click )                

    $PathsGroupBox.Controls.AddRange(@($SharePathLabel, $SharePathTextField, $SharePathLabelSingle, $SharePathSingleTextField, $ShareButton, $CommonButton))

    $CM_form.Controls.AddRange(@($PathsGroupBox))

    #----------------------------------------------------------------------------------
    # Create CM checkboxes GroupBox
    #----------------------------------------------------------------------------------

    $CMGroupBox = New-Object System.Windows.Forms.GroupBox
    $CMGroupBox.location = New-Object System.Drawing.Point($LeftOffset,($TopOffset+20))     # (from left, from top)
    $CMGroupBox.Size = New-Object System.Drawing.Size(240,35)                              # (width, height)

    #----------------------------------------------------------------------------------
    # Create CM Repository Packages Update button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Repo Checkbox' }
    $updateCMCheckbox = New-Object System.Windows.Forms.CheckBox
    $updateCMCheckbox.Text = 'Update SCCM /'
    $updateCMCheckbox.Autosize = $true
    $updateCMCheckbox.Location = New-Object System.Drawing.Point(($LeftOffset-10),($TopOffset-10))

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
        CMTraceLog -Message "updating $($IniFile) setting: ''$($replace)''" -Type $TypeNorm
        (Get-Content $IniFIleFullPath) | Foreach-Object {if ($_ -match $find) {$replace} else {$_}} | Set-Content $IniFIleFullPath

    } # $updateCMCheckbox_Click = 

    $updateCMCheckbox.add_Click($updateCMCheckbox_Click)

    #----------------------------------------------------------------------------------
    # Create CM Distribute Packages button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Repo Checkbox' }
    $CMDistributeheckbox = New-Object System.Windows.Forms.CheckBox
    $CMDistributeheckbox.Text = 'Distribute Packages'
    $CMDistributeheckbox.Autosize = $true
    $CMDistributeheckbox.Location = New-Object System.Drawing.Point(($LeftOffset+95),($TopOffset-10))

    # populate CM Udate checkbox from .INI variable setting - $DistributeCMPackages
    $find = "^[\$]DistributeCMPackages"
    (Get-Content $IniFIleFullPath) | Foreach-Object { 
        if ($_ -match $find) { 
            if ( $_ -match '\$true' ) { 
                $CMDistributeheckbox.Checked = $true 
            } else { 
                $CMDistributeheckbox.Checked = $false 
            }
                        } 
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
        CMTraceLog -Message "updating $($IniFile) setting: ''$($replace)''" -Type $TypeNorm
        (Get-Content $IniFIleFullPath) | Foreach-Object {if ($_ -match $find) {$replace} else {$_}} | Set-Content $IniFIleFullPath

    } # $updateCMCheckbox_Click = 

    $CMDistributeheckbox.add_Click($CMDistributeheckbox_Click)

    $CMGroupBox.Controls.AddRange(@($updateCMCheckbox, $CMDistributeheckbox))
    
    $CM_form.Controls.Add($CMGroupBox)

    #----------------------------------------------------------------------------------
    # Create Models list Checked Grid box - add 1st checkbox column
    # The ListView control allows columns to be used as fields in a row
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating DataGridView' }
    $ListViewWidth = ($FormWidth-80)
    $ListViewHeight = 200
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.location = New-Object System.Drawing.Point(($LeftOffset-10),($TopOffset))
    $dataGridView.height = $ListViewHeight
    $dataGridView.width = $ListViewWidth
    $dataGridView.ColumnHeadersVisible = $true                   # the column names becomes row 0 in the datagrid view
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
    $CheckAll.Top=6
    $CheckAll.Checked = $false

    $CheckAll_Click={

        $state = $CheckAll.Checked
        for ($i = 0; $i -lt $dataGridView.RowCount; $i++) {
            $dataGridView[0,$i].Value = $state  
            $dataGridView.Rows[$i].Cells['Driver'].Value = $state

            if ( $state -eq $false ) {
                foreach ( $cat in $FilterCategories ) {                                   
                    $datagridview.Rows[$i].Cells[$cat].Value = $false                # ... reset categories as well
                }
                get_filters $dataGridView
            } # if ( $state -eq $false )
        }
    } # $CheckAll_Click={

    $CheckAll.add_Click($CheckAll_Click)

    $dataGridView.Controls.Add($CheckAll)
    
    #----------------------------------------------------------------------------------
    # add columns 1, 2 (0 is 1st column)
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'adding SysId, Model columns' }
    $dataGridView.ColumnCount = 3                                # 1st column is checkbox column

    $dataGridView.Columns[1].Name = 'SysId'
    $dataGridView.Columns[1].Width = 40
    $dataGridView.Columns[1].DefaultCellStyle.Alignment = "MiddleCenter"

    $dataGridView.Columns[2].Name = 'Model'
    $dataGridView.Columns[2].Width = 210

    #----------------------------------------------------------------------------------
    # add an 'All' Categories column
    # column 3 (0 is 1st column)
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating All Checkbox column' }
    $CheckBoxesAll = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $CheckBoxesAll.Name = 'All'
    $CheckBoxesAll.width = 28
   
    # $CheckBoxesAll.state = Displayed, Resizable, ResizableSet, Selected, Visible

    #################################################################
    # any click on a checkbox in the column has to also make sure the
    # row is selected
    $CheckBoxesAll_Click = {
        $row = $this.currentRow.index  
        $colClicked = $this.currentCell.ColumnIndex                                       # find out what column was clicked
        $prevState = $dataGridView.rows[$row].Cells[$colClicked].EditedFormattedValue     # value BEFORE click action $true/$false
        $newState = !($prevState)                                                         # value AFTER click action $true/$false

        switch ( $colClicked ) {
            0 { 
                if ( $newState ) {
                    $datagridview.Rows[$row].Cells['Driver'].Value = $newState
                } else {
                    $datagridview.Rows[$row].Cells['All'].Value = $false
                    foreach ( $cat in $FilterCategories ) {                                   
                        $datagridview.Rows[$row].Cells[$cat].Value = $false
                    }
                } # else if ( $newState ) 
            } # 0
            3 {                                                                           # user clicked on 'All' category column
                $datagridview.Rows[$row].Cells[0].Value = $newState                       # ... reset row checkbox as appropriate
                foreach ( $cat in $FilterCategories ) {                                   
                    $datagridview.Rows[$row].Cells[$cat].Value = $newState                # ... reset categories as well
                }
            } # 3
            default {
                foreach ( $cat in $FilterCategories ) {
                    $currColumn = $datagridview.Rows[$row].Cells[$cat].ColumnIndex
                    if ( $colClicked -eq $currColumn ) {
                        continue                                                          # this column already handled by vars above, need to check other categories
                    } else {
                        if ( $datagridview.Rows[$row].Cells[$cat].Value ) {
                            $newState = $true
                        }
                    }
                } # foreach ( $cat in $FilterCategories )
 
                $datagridview.Rows[$row].Cells[0].Value = $newState
            } # default
        } # switch ( $colClicked )
  
    } # $CheckBoxesAll_CellClick

    $dataGridView.Add_Click($CheckBoxesAll_Click)
    #################################################################
   
    [void]$DataGridView.Columns.Add($CheckBoxesAll)

    #----------------------------------------------------------------------------------
    # Add checkbox columns for every category filter
    # column 4 on (0 is 1st column)
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating category columns' }
    foreach ( $id in $FilterCategories ) {
        $temp = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
        $temp.name = $id
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
    # add a repository path as last column
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Repo links column' }
    $LinkColumn = New-Object System.Windows.Forms.DataGridViewColumn
    $LinkColumn.Name = 'Repository'
    $LinkColumn.ReadOnly = $true

    [void]$dataGridView.Columns.Add($LinkColumn,"Repository Path")

    $dataGridView.Columns[8].Width = 200

    #----------------------------------------------------------------------------------
    # next 2 lines clear any selection from the initial data view
    #----------------------------------------------------------------------------------
    $dataGridView.CurrentCell = $dataGridView[1,1]
    $dataGridView.ClearSelection()
    
    ###################################################################################
    # Set initial state for each row      
    
    # uncheck all rows' check column
    for ($i = 0; $i -lt $dataGridView.RowCount; $i++) {
        $dataGridView[0,$i].Value = $false
    }
    ###################################################################################

    #----------------------------------------------------------------------------------
    # Add a grouping box around the Models Grid with its name
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating GroupBox' }

    $CMModlesGroupBox = New-Object System.Windows.Forms.GroupBox
    $CMModlesGroupBox.location = New-Object System.Drawing.Point($LeftOffset,($TopOffset+60))     # (from left, from top)
    $CMModlesGroupBox.Size = New-Object System.Drawing.Size(($ListViewWidth+20),($ListViewHeight+30))       # (width, height)
    $CMModlesGroupBox.text = "HP Models / Repository Filters"

    $CMModlesGroupBox.Controls.AddRange(@($dataGridView))

    $CM_form.Controls.AddRange(@($CMModlesGroupBox))
    
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
    $TextBox.location = New-Object System.Drawing.Point($LeftOffset,($TopOffset+300))            # (from left, from top)
    $TextBox.Size = New-Object System.Drawing.Size(($FormWidth-60),230)             # (width, height)

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
    $RefreshGridButton.Location = New-Object System.Drawing.Point(($LeftOffset+150),($FormHeight-47))    # (from left, from top)
    $RefreshGridButton.Text = 'Refresh Grid Filters'
    $RefreshGridButton.AutoSize=$true

    $RefreshGridButton_Click={
        Get_Filters  $dataGridView
    } # $RefreshGridButton_Click={

    $RefreshGridButton.add_Click($RefreshGridButton_Click)

    $CM_form.Controls.Add($RefreshGridButton)
  
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
    Load_HPModule

    #----------------------------------------------------------------------------------
    # Finally, show the dialog on screen
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'calling get_filters' }
    get_filters $dataGridView

    if ( $DebugMode ) { Write-Host 'calling ShowDialog' }
    $CM_form.ShowDialog() | Out-Null

} # Function CreateForm 

# --------------------------------------------------------------------------
# Start of Script
# --------------------------------------------------------------------------

#CMTraceLog -Message "Starting Script: $scriptName, version $ScriptVersion" -Type $TypeNorm

# Create the GUI and take over all actions, like Report and Download

CreateForm
<#
#>
