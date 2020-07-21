<#
    HP Image Assistant and Softpaq Repository Downloader
#>
param(
	[Parameter(Mandatory = $false,Position = 1,HelpMessage = "Application")]
	[ValidateNotNullOrEmpty()]
	[ValidateSet("sync")]
	$RunMethod = "sync"
)
$ScriptVersion = "1.00 (7/20/2020)"

# get the path to the running script, and populate name of INI configuration file
$scriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

#--------------------------------------------------------------------------------------
$IniFile = "HPIARepo_ini.ps1"                     # assume this INF file in same location as script
$IniFIleFullPath = "$($ScriptPath)\$($IniFile)"

. $IniFIleFullPath                                   # source the code in the INI file      

#--------------------------------------------------------------------------------------
#Script Vars Environment Specific loaded from INI.ps1 file

$CMConnected = $false                                # is a connection to SCCM established?
$SiteCode = $null

#$AdminRights = $false
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

    if ( $Script:DebugMode ) { CMTraceLog -Message "... setting location to: $($SiteCode)" -Type $TypeDebug }
    Set-Location -Path "$($SiteCode):"

    if ( $Script:DebugMode ) { CMTraceLog -Message "... getting CM package: $($lPkgName)" -Type $TypeDebug }
    $lCMRepoPackage = Get-CMPackage -Name $lPkgName -Fast

    if ( $lCMRepoPackage -eq $null ) {
        CMTraceLog -Message "...CM Package missing... Creating New" -Type $TypeNorm
        $lCMRepoPackage = New-CMPackage -Name $lPkgName -Manufacturer "HP"
    }

    #--------------------------------------------------------------------------------
    # update package with info from share folder
    #--------------------------------------------------------------------------------
    if ( $DebugMode ) { CMTraceLog -Message "... setting CM Package Version: $($OSVER), and path: $($pRepoPath)" -Type $TypeDebug }
    #Set-CMPackage -Name $lPkgName -Language $psSoftpaqID
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

    # by now, the Drive or share are available
    if ( $DebugMode ) { CMTraceLog -Message "... checking for Main Repo path" -Type $TypeDebug }

    $lRepoPathSplit = Split-Path -Path $RepoShareMain -Parent        # Get the Path leading to the Main Repo folder
    if ( !(Test-Path $lRepoPathSplit) ) {
        Try {
            New-Item -Path $lRepoPathSplit -ItemType directory
            CMTraceLog -Message "Supporting path created: $($lRepoPathSplit)" -Type $TypeWarn 
        } Catch {
            CMTraceLog -Message "Supporting path creation Failed: $($lRepoPathSplit)" -Type $TypeError
            return
        }
    } # if ( !(Test-Path $lRepoPathSplit) ) 

    #--------------------------------------------------------------------------------
    # Confirm Main Repository folder exists, if not create it

    if ( $DebugMode ) { CMTraceLog -Message "... checking for drive letter" -Type $TypeDebug }

    if ( !(Test-Path $RepoShareMain) ) {

            CMTraceLog -Message '... creating Main Repository Folder = where the Systems individual Repositories will reside' -Type $TypeNorm
            $HPIAShare = New-Item -Path $RepoShareMain -ItemType directory

    } # if ( !(Test-Path $RepoShareMain) )

    #--------------------------------------------------------------------------------
    # now check for each product's repository folder
    # Create the repository folder for each system in the checklist that is selected
    #--------------------------------------------------------------------------------

    if ( $DebugMode ) { CMTraceLog -Message "... checking for existance of selected model's repository path" -Type $TypeDebug }

    for ( $i = 0; $i -lt $pModelsList.RowCount; $i++ ) {

        # it model entry is checked, we need to create a repository
        if ( $i -in $pCheckedItemsList ) {

            $lModelId = $pModelsList[1,$i].Value                                                # column 1 has the Model/Prod ID
            $lModelName = $pModelsList[2,$i].Value                                              # column 2 has the Model name

            $lTempRepoFolder = "$($RepoShareMain)\$($lModelName)"                               # this is the repo folder for this model

            #--------------------------------------------------------------------------------
            # confirm existance, or create repository folder for the current model
            #--------------------------------------------------------------------------------

            CMTraceLog -Message  "--- checking for Repository folder $($lModelName)" -Type $TypeNoNewline

            if ( !(Test-Path $lTempRepoFolder) ) {
                CMTraceLog -Message  "... creating Repository folder" -Type $TypeNorm
                $res = (New-Item -Path $lTempRepoFolder -ItemType directory) | Out-Null

                CMTraceLog -Message  "... initializing repository: $($lTempRepoFolder)" -Type $TypeNorm
                set-location $lTempRepoFolder
                $initOut = (Initialize-Repository) 6>&1
                CMTraceLog -Message  "... Repository Initialization completed: $($Initout)" $TypeWarn 

                CMTraceLog -Message  '... configuring this repository for HP Image Assistant' -Type $TypeNorm
                Set-RepositoryConfiguration -setting OfflineCacheMode -cachevalue Enable 6>&1   # configuring the repo for HP IA's use
            } else {
                CMTraceLog -Message  "... Repository folder exists" -Type $TypeWarn
            } # else if ( $i -in $pCheckedItemsList )

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
            # update filters - every category checked for the current model in this loop
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
    Function get_filters

#>
Function get_filters {
    [CmdletBinding()]
	param( $pModelsList )                             # array of row lines that are checked

    CMTraceLog -Message '[get_filters] Enter' -Type $TypeNorm
    
    Set-Location "C:\" -PassThru

    #--------------------------------------------------------------------------------
    # find out if the share exists, if not, just return

    if ( Test-Path $RepoShareMain ) {
        CMTraceLog -Message '... Main Repo Host Folder exists - will check for selected Model Repositories' -Type $TypeNorm

        #--------------------------------------------------------------------------------
        # now check for each product's repository folder
        # if the repo is created, then check the filters

        for ( $i = 0; $i -lt $pModelsList.RowCount; $i++ ) {

            $lModelId = $pModelsList[1,$i].Value                                                # column 1 has the Model/Prod ID
            $lModelName = $pModelsList[2,$i].Value                                              # column 2 has the Model name

            $lTempRepoFolder = "$($RepoShareMain)\$($lModelName)"                               # this is the repo folder for this model

            # move to location of Repository to use CMSL repo commands
            if ( Test-Path $lTempRepoFolder ) {
                set-location $lTempRepoFolder
                CMTraceLog -Message "... Repository Folder ''$($lTempRepoFolder)'' exists - retrieving filters" -Type $TypeNorm
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
    # Create CM Repository Packages Update button
    if ( $DebugMode ) { Write-Host 'creating Repo Checkbox' }
    $updateCMCheckbox = New-Object System.Windows.Forms.CheckBox
    $updateCMCheckbox.Text = 'Update SCCM Repo Packages'
    $updateCMCheckbox.Autosize = $true
    $updateCMCheckbox.Location = New-Object System.Drawing.Point(($LeftOffset+20),($TopOffset+30))

    # populate CM Udate checkbox from .INI variable setting
    $find = "^[\$]UpdateCMPackages"
    (Get-Content $IniFIleFullPath) | Foreach-Object { 
        if ($_ -match $find) { 
            if ( $_ -match '\$true' ) { 
                $updateCMCheckbox.Checked = $true 
            } else { $updateCMCheckbox.Checked = $false }
                        } 
        } # Foreach-Object

    $Script:UpdateCMPackages = $updateCMCheckbox.Checked

    $updateCMCheckbox_Click = {

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
    # Create 'Debug Mode' - checkmark

    $DebugCheckBox = New-Object System.Windows.Forms.CheckBox
    $DebugCheckBox.Text = 'Debug Mode'
    $DebugCheckBox.UseVisualStyleBackColor = $True
    $DebugCheckBox.location = New-Object System.Drawing.Point(($LeftOffset+240),($TopOffset+30))   # (from left, from top)
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
    # Create OS and OS Version display fields - info from .ini file
    if ( $DebugMode ) { Write-Host 'creating OS Combo and Label' }
    $OSTextLabel = New-Object System.Windows.Forms.Label
    $OSTextLabel.Text = "Win 10 Version:"
    $OSTextLabel.location = New-Object System.Drawing.Point(($LeftOffset+150),($TopOffset+4))    # (from left, from top)
    $OSTextLabel.Size = New-Object System.Drawing.Size(90,25)                               # (width, height)
    #$OSTextField = New-Object System.Windows.Forms.TextBox
    $OSVERComboBox = New-Object System.Windows.Forms.ComboBox
    $OSVERComboBox.Size = New-Object System.Drawing.Size(60,$FieldHeight)                  # (width, height)
    $OSVERComboBox.Location  = New-Object System.Drawing.Point(($LeftOffset+240), ($TopOffset))
    $OSVERComboBox.DropDownStyle = "DropDownList"
    $OSVERComboBox.Name = "OS_Selection"
    $OSVERComboBox.add_MouseHover($ShowHelp)
    
    Foreach ($MenuItem in $OSVALID) {
        [void]$OSVERComboBox.Items.Add($MenuItem);
    }  
    $OSVERComboBox.SelectedItem = $OSVER 

    $CM_form.Controls.AddRange(@($OSTextLabel,$OSVERComboBox))

    #----------------------------------------------------------------------------------
    if ( $DebugMode ) { Write-Host 'creating Share field' }
    $SharePathLabel = New-Object System.Windows.Forms.Label
    $SharePathLabel.Text = "Share:"
    $SharePathLabel.location = New-Object System.Drawing.Point($LeftOffset,($TopOffset+5)) # (from left, from top)
    $SharePathLabel.Size = New-Object System.Drawing.Size(50,20)                            # (width, height)
    $SharePathLabel.TextAlign = "MiddleRight"
    $SharePathTextField = New-Object System.Windows.Forms.TextBox
    $SharePathTextField.Text = "$RepoShareMain"
    $SharePathTextField.Multiline = $false 
    $SharePathTextField.location = New-Object System.Drawing.Point(($LeftOffset+50),($TopOffset+5)) # (from left, from top)
    $SharePathTextField.Size = New-Object System.Drawing.Size(310,$FieldHeight)             # (width, height)
    $SharePathTextField.ReadOnly = $true
    $SharePathTextField.Name = "Share_Path"
    
    $CMGroupBox = New-Object System.Windows.Forms.GroupBox
    $CMGroupBox.location = New-Object System.Drawing.Point(($LeftOffset+360),($TopOffset))     # (from left, from top)
    $CMGroupBox.Size = New-Object System.Drawing.Size(390,60)                              # (width, height)
    $CMGroupBox.text = "Share Path - from ($($IniFile)):"

    $CMGroupBox.Controls.AddRange(@($SharePathLabel, $SharePathTextField))

    $CM_form.Controls.AddRange(@($CMGroupBox))

    #----------------------------------------------------------------------------------
    # Create Models list Checked Grid box - add 1st checkbox column
    # The ListView control allows columns to be used as fields in a row
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
    if ( $DebugMode ) { Write-Host 'creating Repo links column' }
    $LinkColumn = New-Object System.Windows.Forms.DataGridViewColumn
    $LinkColumn.Name = 'Repo'
    $LinkColumn.ReadOnly = $true

    [void]$dataGridView.Columns.Add($LinkColumn,"Repository Path")

    $dataGridView.Columns[8].Width = 200

    #----------------------------------------------------------------------------------
    # next 2 lines clear any selection from the initial data view
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
    if ( $DebugMode ) { Write-Host 'creating GroupBox' }

    $CMModlesGroupBox = New-Object System.Windows.Forms.GroupBox
    $CMModlesGroupBox.location = New-Object System.Drawing.Point($LeftOffset,($TopOffset+60))     # (from left, from top)
    $CMModlesGroupBox.Size = New-Object System.Drawing.Size(($ListViewWidth+20),($ListViewHeight+30))       # (width, height)
    $CMModlesGroupBox.text = "HP Models / Repository Filters"

    $CMModlesGroupBox.Controls.AddRange(@($dataGridView))

    $CM_form.Controls.AddRange(@($CMModlesGroupBox))
    
    #----------------------------------------------------------------------------------
    # Create Output Text Box at the bottom of the dialog
    if ( $DebugMode ) { Write-Host 'creating RichTextBox' }
    $Script:TextBox = New-Object System.Windows.Forms.RichTextBox
    $TextBox.Name = $Script:FormOutTextBox                                          # named so other functions can output to it
    $TextBox.Multiline = $true
    $TextBox.Autosize = $false
    $TextBox.ScrollBars = "Both"
    $TextBox.WordWrap = $true
    $TextBox.location = New-Object System.Drawing.Point($LeftOffset,($TopOffset+300))            # (from left, from top)
    $TextBox.Size = New-Object System.Drawing.Size(($FormWidth-60),230)             # (width, height)

    $CM_form.Controls.AddRange(@($TextBox))

    #----------------------------------------------------------------------------------
    # Add a TextBox wrap CheckBox
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
    # Finally, show the dialog on screen
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
