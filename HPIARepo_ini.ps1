<#
    File: HPIARepo_ini.ps1
    Dan Felman/HP Inc
    10/10/2022
    Updated for Downloader script version >= 1.90
    modify variables as need for use by the Downloader script
    Separated script variables that can mod by Downloader
    Adding support for Win11 w/Downloader v2.00
#>

$FileServerName = $Env:COMPUTERNAME 
#-------------------------------------------------------------------
# set $OS as a read-only script variable with value of 'Win10'
# ... if Windows 11 is now required, modify this entry to 'Win11'
# ... default is set in $v_OS variable

$v_OPSYS = @("Win10", "Win11")
#-------------------------------------------------------------------
$v_OSVALID10 = @("1809", "1909", "2009", "21H1", "21H2","22H2")
$v_OSVALID11 = @("21H2","22H2")

#-------------------------------------------------------------------

# these are the Categories to be selected by HPIA, as needed from the repository
$v_FilterCategories = @('Driver','BIOS', 'Firmware', 'Software')

#-------------------------------------------------------------------
<#
Since version 1.86, script uses these Softpaq names to help populate Platform ID AddOns flag file
... these are just a default list to chose from when adding a model to the product list
... add, remove entries from this list as needed
... these names will be shown in a dialog when adding a model to sync
... >> Add softpaq names (partial names will be matched if possible) to the list, or remove if you do not want to see them as options
#>
$v_Softpaqs = @(
    'Notifications', 'HP Collaboration', 'Presence', 'Tile', 'Power Manager', 'Easy Clean', 'Default Settings', 'MIK', 'Privacy'
)
#-------------------------------------------------------------------

$v_LogFile = "$PSScriptRoot\HPDriverPackDownload.log"          # Default to location where script runs from

# if there errors during startup, set this to $true, for additional info sent to console
$v_DebugMode = $false                                          # Setting managed in GUI (temporarily)

####################################################################################
#
# Settings managed by Script GUI - these are the defaults read by the script at startup
#
####################################################################################

$v_OS = 'Win11'
$v_OSVER = '21H2'
#-------------------------------------------------------------------
<#
 HP Models to be listed in the GUI
 NOTE: Since version 1.90 AddOns list is maintained in a flag file named for the product SysID
       added to the .ADDSOFTWARE folder (example '880D')
 The model table is updated/recreated after a SYNC based on the entries in the GUI
 Only items select for sync get added to this list
#>
$HPModelsTable = @(
	@{ ProdCode = '83B2'; Model = 'HP EliteBook 840 G5 Notebook PC' } 
	@{ ProdCode = '8427'; Model = 'HP ZBook Studio G5 Mobile Workstation' } 
	@{ ProdCode = '8538'; Model = 'HP ProBook 450 G6 Notebook PC' } 
	@{ ProdCode = '8715'; Model = 'HP ProDesk 600 G6 Desktop Mini PC' } 
	@{ ProdCode = '87EA'; Model = 'HP ProBook 640 G8 Notebook PC' } 
	@{ ProdCode = '87ED'; Model = 'HP ProBook 640 G8 Notebook PC' } 
	@{ ProdCode = '81C6'; Model = 'HP Z6 G4 Workstation' }
       )

$v_Continueon404 = $False
$v_KeepFilters = $False

#-------------------------------------------------------------------
# choose/edit a (hardcoded or share) path from next 2 entries for repository locations
# ... this is updated when selecting/creating repository folders

$v_Root_IndividualRepoFolder = "C:\HPIARepo_root"
$v_Root_CommonRepoFolder = "C:\HPIA_Common_11"

$v_CommonRepo = $False

#-------------------------------------------------------------------
# next settings for connecting with Microsoft SCCM/MEM-CM
# they can be modified in the main script, as needed

$v_HPIACMPackage = 'HPIA'            # Package Name for creating/maintaining in SCCM
$v_HPIAVersion = '5.0.2.3827'      # info pulled from Browsing for the app in UI
$v_HPIAPath = '\\CM01\Share\Applications\SP107374 HPIA 5.0.2'            

$v_UpdateCMPackages = $False
$v_DistributeCMPackages = $False
#-------------------------------------------------------------------
