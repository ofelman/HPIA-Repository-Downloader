<#
    HP Image Assistant and Softpaq Repository Downloader
#>
param(
	[Parameter(Mandatory = $false,Position = 1,HelpMessage = "Application")]
	[ValidateNotNullOrEmpty()]
	[ValidateSet("sync")]
	$RunMethod = "sync"
)
$ScriptVersion = "1.10 (7/22/2020)"

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

# if there errors during startup, set this to $true, for additional info produced
$DebugMode = $false
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

} # function Load_HPModule

#=====================================================================================
<#
    Function Test_CMConnection
        The function will test the CM server connection
        and that the Task Sequences required for use of the Script are available in CM
        - will also test that both download and share paths exist
#>
Function Test_CMConnection {

    if ( $Script:CMConnected ) { return $True }                  # already Tested connection

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

    return $boolConnectionRet

} # Function Test_CMConnection

#=====================================================================================
<#
    Function CM_RepoUpdate
#>

Function CM_RepoUpdate {
    [CmdletBinding()]
	param( $pModelName, $pModelProdId, $pRepoPath )                             

    CMTraceLog -Message "[CM_RepoUpdate] Enter" -Type $TypeNorm
    # develop the Package name
    $lPkgName = 'HP-'+$pModelProdId+'-'+$pModelName
    CMTraceLog -Message "... updating SCCM package for repository: $($lPkgName)" -Type $TypeNorm

    if ( $DebugMode ) { CMTraceLog -Message "... setting location to: $($SiteCode)" -Type $TypeDebug }
    Set-Location -Path "$($SiteCode):"

    if ( $DebugMode ) { CMTraceLog -Message "... getting CM package: $($lPkgName)" -Type $TypeDebug }
    $lCMRepoPackage = Get-CMPackage -Name $lPkgName -Fast

    if ( $lCMRepoPackage -eq $null ) {
        CMTraceLog -Message "...CM Package missing... Creating New" -Type $TypeNorm
        $lCMRepoPackage = New-CMPackage -Name $lPkgName -Manufacturer "HP"
    }

    #--------------------------------------------------------------------------------
    # update package with info from share folder
    #--------------------------------------------------------------------------------
    if ( $DebugMode ) { CMTraceLog -Message "... setting CM Package Version: $($OSVER), and path: $($pRepoPath)" -Type $TypeDebug }

	Set-CMPackage -Name $lPkgName -Version "$($OSVER)"
	Set-CMPackage -Name $lPkgName -Path $pRepoPath

    if ( $Script:DebugMode ) { CMTraceLog -Message "... updating CM Distribution Points" -Type $TypeDebug }
    update-CMDistributionPoint -PackageId $lCMRepoPackage.PackageID

    $lCMRepoPackage = Get-CMPackage -Name $lPkgName -Fast                               # make suer we are woring with updated package

    Set-Location -Path "C:"
    CMTraceLog -Message "[CM_RepoUpdate] Done" -Type $TypeNorm

} # Function CM_RepoUpdate

#=====================================================================================
<#
    Function init_repoFolder
        This function will create a repository foldern and initialize it for HPIA, if necessary
        Args:
            $pRepoFOlder: folder to validate, or create
            $pInitRepository: variable - if $true, initialize repository, otherwise, ignore
#>
Function init_repoFolder {
    [CmdletBinding()]
	param( $pRepoFolder,
            $pInitRepository )

    CMTraceLog -Message "[init_repoFolder] Enter" -Type $TypeNorm

    $retRepoCreated = $false

    if ( (Test-Path $pRepoFolder) -and ($pInitRepository -eq $false) ) {
        $retRepoCreated = $true
    } else {
        $pCurrentLoc = Get-Location

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
                CMTraceLog -Message "[init_repoFolder] Done" -Type $TypeNorm
                return $retRepoCreated
            } # Catch
        } # if ( !(Test-Path $lRepoPathSplit) ) 

        #--------------------------------------------------------------------------------
        # now add the Repo folder if it doesn't exist

        if ( !(Test-Path $pRepoFolder) ) {
            CMTraceLog -Message '... creating Repository Folder' -Type $TypeNorm
            New-Item -Path $pRepoFolder -ItemType directory
        } # if ( !(Test-Path $RepoShareMain) )

        $retRepoCreated = $true

        #--------------------------------------------------------------------------------
        # if needed, check on repository to initialize

        if ( $pInitRepository -and !(test-path "$pRepoFolder\.Repository")) {
            Set-Location $pRepoFolder
            $initOut = (Initialize-Repository) 6>&1
            CMTraceLog -Message  "... Repository Initialization done: $($Initout)"  -Type $TypeNorm 

            CMTraceLog -Message  '... configuring this repository for HP Image Assistant' -Type $TypeNorm
            Set-RepositoryConfiguration -setting OfflineCacheMode -cachevalue Enable 6>&1   # configuring the repo for HP IA's use

        } # if ( $pInitRepository -and !(test-path "$pRepoFOlder\.Repository"))

        Set-Location $pCurrentLoc
    } #

    CMTraceLog -Message "[init_repoFolder] Done" -Type $TypeNorm
    return $retRepoCreated

} # Function init_repoFolder

#=====================================================================================
<#
    Function invoke_Sync
        This function is designed to do the following:
            - create repository folders (if not yet available) for every HP Model in the list
            - initialize each folder as a SOftpaq Repository for use by HP Image Assistant
            - reset the CMSL filters based on the user's GUI input
            - invoke a repository sync on every model folder (NOTE: will result in multiple duplicates)
            - invoke a repository cleanup on every model folder - to remove superseeded softpaqs

    expects parameter 
        - Gui Model list table
        - list of checked entries from the table checkmarks
        - boolean downloadAndUpdate - for future use
#>
Function invoke_Sync {
    [CmdletBinding()]
	param( $pModelsList,                                             # array of row lines that are checked
            $pCheckedItemsList)                                      # array of rows selected
    
    CMTraceLog -Message "[invoke_Sync] Enter" -Type $TypeNorm

    if ( $DebugMode ) { CMTraceLog -Message "... for OS version: $($OSVER)" -Type $TypeDebug }
    
    Set-Location "C:"

    #--------------------------------------------------------------------------------
    # Confirm path to Main Repository folder exists, if not create it
    # ... but first check to make sure the Drive the path is on is available (-qualifier option)

    # are we dealing with a Drive or a share? if a drive, is it accessible?

    if ( $DebugMode ) { CMTraceLog -Message "... checking for drive letter" -Type $TypeDebug }

    if ( $RepoShareMain.Substring(1,1) -eq ':' ) {
        $lDriveLetter = Split-Path $RepoShareMain -Qualifier
        if ( $DebugMode ) { CMTraceLog -Message "... drive letter $($lDriveLetter)" -Type $TypeDebug }
        if ( !(Test-Path $lDriveLetter) ) {
            CMTraceLog -Message "[invoke_Sync] Exit: Supporting path creation Failed on drive: $($lDriveLetter)" -Type $TypeError
            return
        }
    } else {
        if ( $DebugMode ) { CMTraceLog -Message "... NO drive letter in path" -Type $TypeDebug }
    } # else if ( $RepoShareMain.Substring(1,1) -eq ':' )

    #--------------------------------------------------------------------------------
    # by now, we checked and the Drive or share are available, so make sure the folder is there
    if ( $DebugMode ) { CMTraceLog -Message "... checking for Main Repo path" -Type $TypeDebug }
        
    if ( $Script:singleRepo -eq $false ) {
        init_repoFolder $RepoShareMain $false                  # make sure Main repo folder exists, or create it - no init
    } else {
        init_repoFolder $RepoShareMain $true                   # check folder, and initialize for HPIA
    }    

    #--------------------------------------------------------------------------------
    # now check for each product's repository folder
    # Create the repository folder for each system in the checklist that is selected
    #--------------------------------------------------------------------------------

    if ( $DebugMode ) { CMTraceLog -Message "... checking for existance of selected model's repository path" -Type $TypeDebug }

    # go through every selected HP Model in the list

    for ( $i = 0; $i -lt $pModelsList.RowCount; $i++ ) {

        # if model entry is checked, we need to create a repository, but ONLY if $singleRepo = $false
        if ( $i -in $pCheckedItemsList ) {

            $lModelId = $pModelsList[1,$i].Value                                                # column 1 has the Model/Prod ID
            $lModelName = $pModelsList[2,$i].Value                                              # column 2 has the Model name

            if ( $Script:singleRepo -eq $false ) {
                $lTempRepoFolder = "$($RepoShareMain)\$($lModelName)"                               # this is the repo folder for this model
                #--------------------------------------------------------------------------------
                # confirm existance, or create repository folder for the current model
                #--------------------------------------------------------------------------------
                CMTraceLog -Message  "--- checking for Repository folder ''$($lTempRepoFolder)''" -Type $TypeNorm
                init_repoFolder $lTempRepoFolder $true
            } else {
                $lTempRepoFolder = $RepoShareMain
            } # else if ( $Script:singleRepo -eq $false )
 
            # move to location of Repository to use CMSL repo commands
            set-location $lTempRepoFolder
            
            #--------------------------------------------------------------------------------
            # clean up filters for the current model in this loop
            #--------------------------------------------------------------------------------

            CMTraceLog -Message  "... removing filters for Sysid: $($lModelName)/$($lModelId)"  -Type $TypeNorm
            $lres = (Remove-RepositoryFilter -platform $lModelId -yes 6>&1)      
            if ( $debugMode ) {
                CMTraceLog -Message "... removed filter: $($lres)" -Type $TypeWarn
            }

            #--------------------------------------------------------------------------------
            # update filters - every category checked for the current model in this for loop
            #--------------------------------------------------------------------------------

            if ( $DebugMode ) { CMTraceLog -Message "... updating grid with Model's repository filters" -Type $TypeDebug }

            foreach ( $cat in $FilterCategories ) {

                if ( $datagridview.Rows[$i].Cells[$cat].Value ) {
                    CMTraceLog -Message  "... adding filter: -platform $($lModelId) -os win10 -osver $OSVER -category $($cat) -characteristic ssm" -Type $TypeNoNewline
                    $lRes = (Add-RepositoryFilter -platform $lModelId -os win10 -osver $OSVER -category $cat -characteristic ssm 6>&1)
                    CMTraceLog -Message $lRes -Type $TypeWarn 
                }
            } # foreach ( $cat in $FilterCategories )

            #--------------------------------------------------------------------------------
            # show repository path for this model (model is checked) (col 8 is links col)
            #--------------------------------------------------------------------------------

            $datagridview[8,$i].Value = $lTempRepoFolder

            #--------------------------------------------------------------------------------
            # update repository softpaqs with sync command and then cleanup
            #--------------------------------------------------------------------------------

            CMTraceLog -Message  '... invoking repository sync - please wait !!!' -Type $TypeNorm
            $lRes = (invoke-repositorysync 6>&1)
            CMTraceLog -Message "   ... $($lRes)" -Type $TypeWarn 

            CMTraceLog -Message  '... invoking repository cleanup ' -Type $TypeNoNewline 
            $lRes = (invoke-RepositoryCleanup 6>&1)
            CMTraceLog -Message "... $($lRes)" -Type $TypeWarn

            #--------------------------------------------------------------------------------
            # update SCCM Repository package, if user allows
            #--------------------------------------------------------------------------------

            if ( $Script:UpdateCMPackages ) {
                CM_RepoUpdate $lModelName $lModelId $lTempRepoFolder
            }

        } # if ( $i -in $pCheckedItemsList )

     } # for ( $i = 0; $i -lt $pModelsList.RowCount; $i++ )

    #--------------------------------------------------------------------------------

    CMTraceLog -Message "[invoke_Sync] Done" -Type $TypeNorm

} # Function invoke_Sync

#=====================================================================================
<#
    Function clear_datagrid
        clear all checkmarks and last path column '8'

#>
Function clear_datagrid {
    [CmdletBinding()]
	param( $pDataGrid )                             

    CMTraceLog -Message '[clear_datagrid] Enter' -Type $TypeNorm

    for ( $row = 0 ; $row -lt $pDataGrid.RowCount ; $row++ ) {
        for ( $col = 0 ; $col -lt $pDataGrid.ColumnCount ; $col++ ) {
            if ( $col -in @(0,3,4,5,6,7) ) {
                $pDataGrid[$col,$row].value = $false
            } else {
                if ( $col -in @(8) ) {
                    $pDataGrid[$col,$row].value = ''
                }
            }
        }
    } # for ( $row = 0 ; $row -lt $pDataGrid.RowCount ; $row++ ) 

    CMTraceLog -Message '[clear_datagrid] Done' -Type $TypeNorm
} # Function clear_datagrid

#=====================================================================================
<#
    Function get_filters

#>
Function get_filters {
    [CmdletBinding()]
	param( $pModelsList )                             # array of row lines that are checked

    CMTraceLog -Message '[get_filters] Enter' -Type $TypeNorm
    
    Set-Location "C:\" #-PassThru

    #--------------------------------------------------------------------------------
    # find out if the share exists, if not, just return
    #--------------------------------------------------------------------------------
    if ( Test-Path $RepoShareMain ) {
        CMTraceLog -Message '... Main Repo Host Folder exists - will check for selected Models' -Type $TypeNorm

        if ( $Script:singleRepo ) {
            # see if the repository was configured for HPIA
            set-location $RepoShareMain
            if ( !(Test-Path "$($RepoShareMain)\.Repository") ) {
                CMTraceLog -Message "... Repository Folder not initialized" -Type $TypeNorm
                clear_DataGrid $datagridview
                CMTraceLog -Message '[get_filters] Done' -Type $TypeNorm
                return
            } 
        } # if ( $Script:singleRepo )

        #--------------------------------------------------------------------------------
        # now check for each product's repository folder
        # if the repo is created, then check the filters
        #--------------------------------------------------------------------------------
        for ( $i = 0; $i -lt $pModelsList.RowCount; $i++ ) {

            $lModelId = $pModelsList[1,$i].Value                                                # column 1 has the Model/Prod ID
            $lModelName = $pModelsList[2,$i].Value                                              # column 2 has the Model name

            if ( $Script:singleRepo ) {
                $lTempRepoFolder = $RepoShareMain
            } else {
                $lTempRepoFolder = "$($RepoShareMain)\$($lModelName)"                               # this is the repo folder for this model
            } # else if ( $Script:singleRepo )

            # move to location of Repository to use CMSL repo commands
            if ( Test-Path $lTempRepoFolder ) {
                set-location $lTempRepoFolder
                CMTraceLog -Message "... Repository Folder ''$($lTempRepoFolder)'' for ''$($lModelName)'' exists - retrieving filters" -Type $TypeNorm
                ###### filters sample obtained with get-repositoryinfo: 
                # platform        : 8438
                # operatingSystem : win10:2004 win10:2004
                # category        : BIOS firmware
                # releaseType     : *
                # characteristic  : ssm
                ###### 
                $lProdFilters = (get-repositoryinfo).Filters
        
                if ( $lProdFilters ) {
                    CMTraceLog -Message "... populating filter categories $($lTempRepoFolder)" -Type $TypeDebug
                    $pModelsList[0,$i].Value = $true
                    foreach ( $cat in  ($lProdFilters.category.split(' ')) ) {
                        $datagridview.Rows[$i].Cells[$cat].Value = $true
                    }
                    #--------------------------------------------------------------------------------
                    # show repository path for this model (model is checked) (col 8 is links col)
                    #--------------------------------------------------------------------------------

                    $datagridview[8,$i].Value = $lTempRepoFolder

                } # if ( $lProdFilters )
            } # if ( Test-Path $lTempRepoFolder )
        } # for ( $i = 0; $i -lt $pModelsList.RowCount; $i++ ) 

    } else {
        CMTraceLog -Message '... Main Repo Host Folder ''$($RepoShareMain)''does NOT exist' -Type $TypeNorm
    } # else if ( !(Test-Path $RepoShareMain) )

    CMTraceLog -Message '[get_filters] Done' -Type $TypeNorm

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
    $buttonSync.Location = New-Object System.Drawing.Point(($LeftOffset+20),($TopOffset-1))

    $buttonSync.add_click( {

        # Modify INI file with newly selected OSVER
        #####################################################################
        if ( $Script:OSVER -ne $OSVERComboBox.Text ) {
            if ( $DebugMode ) { CMTraceLog -Message 'modifying INI file with selected OSVER' -Type $TypeNorm }
            $find = "^[\$]OSVER"
            $replace = "`$OSVER = ""$($OSVERComboBox.Text)"""  
            (Get-Content $IniFIleFullPath) | Foreach-Object {if ($_ -match $find) {$replace} else {$_}} | Set-Content $IniFIleFullPath
        } 
        #####################################################################
        $Script:OSVER = $OSVERComboBox.Text
        if ( $Script:OSVER -in $Script:OSVALID ) {
            # get a list of all checked entries (row numbers, starting with 0)
            $lCheckedListArray = @()
            for ($i = 0; $i -lt $dataGridView.RowCount; $i++) {
                if ($dataGridView[0,$i].Value) {
                    $lCheckedListArray += $i
                } # if  
            } # for
            if ( $updateCMCheckbox.checked ) {
                if ( ($Script:CMConnected = Test_CMConnection) ) {
                    if ( $DebugMode ) { Write-Host 'Script connected to CM' } 
                }
            }
            invoke_Sync $dataGridView $lCheckedListArray

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
    $OSTextLabel.location = New-Object System.Drawing.Point(($LeftOffset+90),($TopOffset+4))    # (from left, from top)
    $OSTextLabel.Size = New-Object System.Drawing.Size(90,25)                               # (width, height)
    #$OSTextField = New-Object System.Windows.Forms.TextBox
    $OSVERComboBox = New-Object System.Windows.Forms.ComboBox
    $OSVERComboBox.Size = New-Object System.Drawing.Size(60,$FieldHeight)                  # (width, height)
    $OSVERComboBox.Location  = New-Object System.Drawing.Point(($LeftOffset+180), ($TopOffset))
    $OSVERComboBox.DropDownStyle = "DropDownList"
    $OSVERComboBox.Name = "OS_Selection"
    $OSVERComboBox.add_MouseHover($ShowHelp)
    
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
    $DebugCheckBox.location = New-Object System.Drawing.Point(($LeftOffset+260),$TopOffset)   # (from left, from top)
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
    # Create Single Repository checkbox
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Single Repo Checkbox' }
    $singleRepoheckbox = New-Object System.Windows.Forms.CheckBox
    $singleRepoheckbox.Text = 'Use Single Repository'
    $singleRepoheckbox.Autosize = $true
    $singleRepoheckbox.Location = New-Object System.Drawing.Point(($LeftOffset+20),($TopOffset+30))

    # populate CM Udate checkbox from .INI variable setting
    $find = "^[\$]SingleRepo"
    (Get-Content $IniFIleFullPath) | Foreach-Object { 
        if ($_ -match $find) { 
            if ( $_ -match '\$true' ) { 
                $singleRepoheckbox.Checked = $true                          # set the visual default from the INI setting
            } else { 
                $singleRepoheckbox.Checked = $false 
            }
                        } 
        } # Foreach-Object

    $Script:SingleRepo  = $singleRepoheckbox.Checked                       

    $singleRepoheckbox_Click = {
        $find = "^[\$]SingleRepo"
        if ( $singleRepoheckbox.checked ) {
            $Script:SingleRepo = $true
        } else {
            $Script:SingleRepo = $false
        }
        get_filters $dataGridView                                          # re-populate categories for available systems
        $replace = "`$SingleRepo = `$$Script:SingleRepo"                   # set up the replacing string to either $false or $true from ini file
        CMTraceLog -Message "updating INI setting to ''$($replace)''" -Type $TypeNorm
        (Get-Content $IniFIleFullPath) | Foreach-Object {if ($_ -match $find) {$replace} else {$_}} | Set-Content $IniFIleFullPath

    } # $updateCMCheckbox_Click = 

    $singleRepoheckbox.add_Click($singleRepoheckbox_Click)
    $CM_form.Controls.Add($singleRepoheckbox)

    #----------------------------------------------------------------------------------
    # Create CM Repository Packages Update button
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Repo Checkbox' }
    $updateCMCheckbox = New-Object System.Windows.Forms.CheckBox
    $updateCMCheckbox.Text = 'Update SCCM Repo Packages'
    $updateCMCheckbox.Autosize = $true
    $updateCMCheckbox.Location = New-Object System.Drawing.Point(($LeftOffset+190),($TopOffset+30))

    # populate CM Udate checkbox from .INI variable setting - $UpdateCMPackages
    $find = "^[\$]UpdateCMPackages"
    (Get-Content $IniFIleFullPath) | Foreach-Object { 
        if ($_ -match $find) { 
            if ( $_ -match '\$true' ) { 
                $updateCMCheckbox.Checked = $true 
            } else { 
                $updateCMCheckbox.Checked = $false 
            }
                        } 
        } # Foreach-Object

    $Script:UpdateCMPackages = $updateCMCheckbox.Checked

    $updateCMCheckbox_Click = {
        $find = "^[\$]UpdateCMPackages"
        if ( $updateCMCheckbox.checked ) {
            $Script:UpdateCMPackages = $true
        } else {
            $Script:UpdateCMPackages = $false
        }
        $replace = "`$UpdateCMPackages = `$$Script:UpdateCMPackages"                   # set up the replacing string to either $false or $true from ini file
        CMTraceLog -Message "updating INI setting to ''$($replace)''" -Type $TypeNorm
        (Get-Content $IniFIleFullPath) | Foreach-Object {if ($_ -match $find) {$replace} else {$_}} | Set-Content $IniFIleFullPath

    } # $updateCMCheckbox_Click = 

    $updateCMCheckbox.add_Click($updateCMCheckbox_Click)

    $CM_form.Controls.Add($updateCMCheckbox)

    #----------------------------------------------------------------------------------
    # add share info field
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Share field' }
    $SharePathLabel = New-Object System.Windows.Forms.Label
    $SharePathLabel.Text = "Share:"
    $SharePathLabel.location = New-Object System.Drawing.Point(($LeftOffset-10),($TopOffset+5)) # (from left, from top)
    $SharePathLabel.Size = New-Object System.Drawing.Size(50,20)                            # (width, height)
    $SharePathLabel.TextAlign = "MiddleRight"
    $SharePathTextField = New-Object System.Windows.Forms.TextBox
    $SharePathTextField.Text = "$RepoShareMain"
    $SharePathTextField.Multiline = $false 
    $SharePathTextField.location = New-Object System.Drawing.Point(($LeftOffset+50),($TopOffset+5)) # (from left, from top)
    $SharePathTextField.Size = New-Object System.Drawing.Size(290,$FieldHeight)             # (width, height)
    $SharePathTextField.ReadOnly = $true
    $SharePathTextField.Name = "Share_Path"
    
    $CMGroupBox = New-Object System.Windows.Forms.GroupBox
    $CMGroupBox.location = New-Object System.Drawing.Point(($LeftOffset+370),($TopOffset))     # (from left, from top)
    $CMGroupBox.Size = New-Object System.Drawing.Size(370,60)                              # (width, height)
    $CMGroupBox.text = "Share Path - from ($($IniFile)):"

    $CMGroupBox.Controls.AddRange(@($SharePathLabel, $SharePathTextField))

    $CM_form.Controls.AddRange(@($CMGroupBox))

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
            $datagridview.Rows[$i].Cells['Driver'].Value = $state
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
    $LinkColumn.Name = 'Repo'
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
    # Add a TextBox wrap CheckBox
    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Word Wrap button' }
    $CheckWordwrap = New-Object System.Windows.Forms.CheckBox
    $CheckWordwrap.Location = New-Object System.Drawing.Point($LeftOffset,($FormHeight-50))    # (from left, from top)
    $CheckWordwrap.Text = 'Wrap Text'
    $CheckWordwrap.AutoSize=$true
    $CheckWordwrap.Checked = $false

    $CheckWordwrap_Click={
        $state = $CheckWordwrap.Checked
        if ( $CheckWordwrap.Checked ) {
            $TextBox.WordWrap = $true
        } else {
            $TextBox.WordWrap = $false
        }
    } # $CheckWordwrap_Click={

    $CheckWordwrap.add_Click($CheckWordwrap_Click)

    $CM_form.Controls.Add($CheckWordwrap)
 
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
