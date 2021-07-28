<#
    File: HPIARepo_ini.ps1

    Dan Felman/HP Inc
    07/27/2021
    Updated for Downloader script version >= 1.76

    modify variables as need for use by the Downloader script
    Separated script variables that can ne mod by Downloader
#>

# set $OS as a read-only script variable with value of 'Win10'
if ( Test-Path variable:OS) {
    Set-Variable -Name OS -value "Win10" -Option ReadOnly -ErrorAction SilentlyContinue
} else {
    New-Variable -Name OS -value "Win10" -Option ReadOnly
}
$v_OSVALID = @("1809", "1909", "2004", "2009", "2104")

# these are the Categories to be selected by HP IA, as needed from the repository
$v_FilterCategories = @('Driver','BIOS', 'Firmware', 'Software')

#-------------------------------------------------------------------
#
# NEW 1.30 - added list of individual softpaq names to maintain in repositories
# ... if model table recreated in GUI, list will need to be added to each needed entry
# 
$v_Softpaqs = @(
    'Hotkeys', 'Notifications' , 'Presence', 'Tile', 'Power Manager', 'Easy Clean', 'Default Settings'
)
# EXAMPLE: 

# { ProdCode = '8723'; Model = 'HP ZBook Firefly 14 G7 Mobile Workstation' ; AddOns = 'Hotkeys', 'Notifications' }

# HP Models to be listed in the GUI
# NOTE: field 'SqName' is optional in the model list, and contain additional HP Softpaqs to be added
#     and maintained by the script for HPIA to find

$HPModelsTable = @(
	@{ ProdCode = '83B2'; Model = 'HP EliteBook 840 G5 Notebook PC' ; AddOns = 'Hotkeys', 'Notifications' } 
	@{ ProdCode = '876D'; Model = 'HP EliteBook x360 1030 G7 Notebook PC' ; AddOns = 'Hotkeys', 'Notifications' } 
	@{ ProdCode = '83EE'; Model = 'HP ProDesk 600 G4 Small Form Factor PC' } 
	@{ ProdCode = '842A'; Model = 'HP ZBook 15 G5 Mobile Workstation' ; AddOns = 'Hotkeys', 'Notifications' } 
	@{ ProdCode = '8723'; Model = 'HP ZBook Firefly 14 G7 Mobile Workstation' ; AddOns = 'Hotkeys', 'Notifications' }
	)

#-------------------------------------------------------------------
$v_LogFile = "$PSScriptRoot\HPDriverPackDownload.log"          # Default to location where script runs from

# if there errors during startup, set this to $true, for additional info
$v_DebugMode = $false                                          # Setting managed in GUI (temporarily)

####################################################################################
#
# Settings managed by Script GUI - these are the defaults read by the script
#
####################################################################################

$v_OSVER = "2009"

$v_Continueon404 = $False
$v_KeepFilters = $False               # 7/31 NEW: setting to keep filters (script version 1.35)

#-------------------------------------------------------------------
# choose/edit a (hardcoded or share) path from next 2 entries for repository locations

$FileServer = $Env:COMPUTERNAME 

# Folder to use for multiple repositories/one per model
$v_Root_IndividualRepoFolder = "C:\HPIA_Repo_Head"

# and this for use when selecting a single repository 
$v_Root_CommonRepository = "C:\HPIACommonRepository"

$v_CommonRepo = $True

#-------------------------------------------------------------------
# next settings for connecting with Microsoft SCCM/MECM
# they can be modified in the main script, as needed

$HPIACMPackage = 'HPIA'             # Package Name for creating/maintaining in SCCM
$v_HPIAVersion = '5.0.2.3827'         # info pulled from Browsing for the app in UI
$v_HPIAPath = '\\CM01\Share\Applications\SP107374 HPIA 5.0.2'            

$v_UpdateCMPackages = $False
$v_DistributeCMPackages = $False
#-------------------------------------------------------------------
