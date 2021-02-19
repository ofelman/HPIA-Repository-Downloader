<#
    File: HPIARepo_ini.ps1

    Dan Felman/HP Inc
    10/20/2020
    2/19/2021 - Update

    modify variables as need for use by the Downloader script
    Separated script variables that can ne mod by Downloader
#>

if ( Test-Path variable:OS) {
    Set-Variable -Name OS -value "Win10" -Option ReadOnly -ErrorAction SilentlyContinue
} else {
    New-Variable -Name OS -value "Win10" -Option ReadOnly
}
$OSVALID = @("1909", "2004", "2009")  # add OS version entries as needed to this variable

# these are the Categories to be selected by HP IA - added as grid columns
$FilterCategories = @('Driver','BIOS', 'Firmware', 'Software')

#-------------------------------------------------------------------
# These are sample Softpaq lists that we will download for a model
#
# NEW 1.30 - added list of individual softpaq names to host in repository
# ***** EXAMPLES *****

$NBSet1 = 'HP Collaboration Keyboard Software' , 'HP Hotkey Support' , 'System Default Settings for Windows 10' , 'HP Notifications'
$NBSet2 = 'HP Hotkey Support' , 'System Default Settings for Windows 10' , 'HP Notifications'
$NBSetG7 = 'HP Hotkey Support' , 'HP Notifications', '107683'                   # sp107683='HP Programmable Key (SA) - version 1.0.13'
$NBSet1KG7 = 'HP Hotkey Support' , 'HP Notifications', 'Presence', '107683'     # sp107683='HP Programmable Key (SA) - version 1.0.13'
$DTSet1 = 'HP Notifications'

# example HP Models to be listed in the GUI
# variable 'SqName' is optional in the model list
# NEW: versoin 1.75 adds ability to modify this list from existing, imported, repository

$HPModelsTable = @(
	@{ ProdCode = '83EE'; Model = 'HP ProDesk 600 G4 Small Form Factor PC' } 
	@{ ProdCode = '83B2'; Model = 'HP EliteBook 846 G5 Notebook PC' } 
	@{ ProdCode = '8723'; Model = 'HP ZBook Firefly 14 G7 Mobile Workstation' }
	)

#-------------------------------------------------------------------
# choose/edit a (remote or local) path from next 2 entries for repository locations
# ... examples show both a \\ share and a drive letter entries ...
$FileServerName = $Env:COMPUTERNAME 

# next for use in multiple repositories/one per model
$RepoShareMain = $RepoShareMain = "\\$($FileServerName)\share\softpaqs\HPIARepoHead"
#$RepoShareMain = "C:\share\softpaqs\HPIARepoHead"

# and this for use when selecting a single repository (use with Script version >= 1.15
#$RepoShareCommon = "\\$($FileServerName)\share\softpaqs\HPIACommonRepository"
$RepoShareCommon = "C:\HPIACommonRepository"

#-------------------------------------------------------------------

$LogFile = "$PSScriptRoot\HPDriverPackDownload.log"      # Default to location where script runs from


$DebugMode = $false              # This allows startup debug info to be displayed, prior to GUI form creation                 

####################################################################################
#
# Settings managed by Script GUI - these are the defaults read by the script
#
####################################################################################

$OSVER = "1909"                  # this is the default version to be used

#-------------------------------------------------------------------
# Items related to HPIA AND SCCM//MECM connection
#-------------------------------------------------------------------
$HPIACMPackage = 'HPIA'             # Package Name for creating/maintaining in SCCM
$HPIAVersion = '5.0.2.3827'         # info pulled from Browse app in UI
$HPIAPath = '\\CM01\Share\Applications\SP107374 HPIA 5.0.2'            
#-------------------------------------------------------------------

$UpdateCMPackages = $False          
$Continueon404 = $false             # useful during CMSL sync process
$KeepFilters = $False               # 7/31 NEW: setting to keep filters (script version 1.35)
$CommonRepo = $True
