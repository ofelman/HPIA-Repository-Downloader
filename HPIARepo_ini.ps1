<#
    File: HPIARepo_ini.ps1
    Dan Felman/HP Inc
    08/10/2021
    Updated for Downloader script version >= 1.76
    modify variables as need for use by the Downloader script
    Separated script variables that can mod by Downloader
#>

$FileServerName = $Env:COMPUTERNAME 
#-------------------------------------------------------------------
# set $OS as a read-only script variable with value of 'Win10'
if ( Test-Path variable:OS) {
    Set-Variable -Name OS -value "Win10" -Option ReadOnly -ErrorAction SilentlyContinue
} else {
    New-Variable -Name OS -value "Win10" -Option ReadOnly
}
#-------------------------------------------------------------------
$v_OSVALID = @("1809", "1909", "2009", "21H1", "21H2")

# these are the Categories to be selected by HP IA, as needed from the repository
$v_FilterCategories = @('Driver','BIOS', 'Firmware', 'Software')

<#-------------------------------------------------------------------
Since version 1.86, script uses these Softpaq names to help populate Platfor ID AddOns flag file
... add, remove any as needed to this list
... these names will be shown in a dialog when adding a model to the display list to sync
#>
$v_Softpaqs = @(
    'Notifications', 'HP Collaboration Keyboard Software', 'Presence', 'Tile', 'Power Manager', 'Easy Clean', 'Default Settings', 'MIK', 'Privacy'
)
#-------------------------------------------------------------------

# HP Models to be listed in the GUI
# NOTE: field AddOns is optional in the model list, and contain additional HP Softpaqs to be added
#     and maintained by the script for HPIA to find

$HPModelsTable = @(
	@{ ProdCode = '80FC'; Model = 'HP Elite x2 1012 G1 Tablet' } 
	@{ ProdCode = '8438'; Model = 'HP EliteBook x360 1030 G3 Notebook PC' }
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

$v_OSVER = "2009"

$v_Continueon404 = $False
$v_KeepFilters = $False

#-------------------------------------------------------------------
# choose/edit a (hardcoded or share) path from next 2 entries for repository locations

$v_Root_IndividualRepoFolder = "C:\HPIA_Repo_Head"
$v_Root_CommonRepoFolder = "C:\HPIA_NEW3"

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
