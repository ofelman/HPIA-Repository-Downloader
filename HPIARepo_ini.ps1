<#
    File: HPIARepo_ini.ps1

    Dan Felman/HP Inc
    10/20/2020

    modify variables as need for use by the Downloader script
    Separated script variables that can ne mod by Downloader
#>

if ( Test-Path variable:OS) {
    Set-Variable -Name OS -value "Win10" -Option ReadOnly -ErrorAction SilentlyContinue
} else {
    New-Variable -Name OS -value "Win10" -Option ReadOnly
}
$OSVALID = @("1809", "1903", "1909", "2004")

# these are the Categories to be selected by HP IA, as needed from the repository
$FilterCategories = @('Driver','BIOS', 'Firmware', 'Software')

#-------------------------------------------------------------------
# These are sample Softpaq lists that we will download for a model
#
# NEW 1.30 - added list of individual softpaq names to host in repository
#
$NBSet1 = 'HP Collaboration Keyboard Software' , 'HP Hotkey Support' , 'System Default Settings for Windows 10' , 'HP Notifications'
$NBSet2 = 'HP Hotkey Support' , 'System Default Settings for Windows 10' , 'HP Notifications'
$NBSetG7 = 'HP Hotkey Support' , 'HP Notifications', '107683'                   # sp107683='HP Programmable Key (SA) - version 1.0.13'
$NBSet1KG7 = 'HP Hotkey Support' , 'HP Notifications', 'Presence', '107683'     # sp107683='HP Programmable Key (SA) - version 1.0.13'
$DTSet1 = 'HP Notifications'

# example HP Models to be listed in the GUI
# variable 'SqName' is optional in the model list

$HPModelsTable = @(
	@{ ProdCode = '805D'; Model = "HP ProDesk 600 G2";                                        }
    @{ ProdCode = '83EE'; Model = "HP ProDesk 600 G4 SFF";                                    }
	@{ ProdCode = '8593'; Model = "HP EliteDesk 800 G5 Desktop Mini PC";    SqName = $DTSet1  }
    @{ ProdCode = '8549'; Model = "HP EliteBook 840 G6 Healthcare Edition"; SqName = $NBSet1  }
    @{ ProdCode = '859F'; Model = "HP EliteOne 800 G5 All-in-One";          SqName = $DTSet1  }
    @{ ProdCode = '83B3'; Model = "HP ELITEBOOK 830 G5";                    SqName = $NBSet1  }
	@{ ProdCode = '83B2'; Model = "HP ELITEBOOK 840 G5";                    SqName = $NBSet1  }
	@{ ProdCode = '8549'; Model = "HP ELITEBOOK 840 G6";                    SqName = $NBSet2  }
	@{ ProdCode = '8723'; Model = "HP ELITEBOOK 840 G7";                    SqName = $NBSetG7 }
	@{ ProdCode = '8470'; Model = "HP ELITEBOOK X360 1040 G5";              SqName = $NBSet1  }
	@{ ProdCode = '80D4'; Model = "HP ZBook Studio G3"                                        }	
	@{ ProdCode = '844F'; Model = "HP ZBook Studio x360 G5";                SqName = $NBSet1  }
    @{ ProdCode = '844F'; Model = "HP ZBook Studio G5";                                       }
	)

#-------------------------------------------------------------------
# choose/edit a (remote or local) path from next 2 entries for repository locations

$FileServerName = $Env:COMPUTERNAME 

# next for use in multiple repositories/one per model
$RepoShareMain = $RepoShareMain = "\\$($FileServerName)\share\softpaqs\HPIARepoHead"
#$RepoShareMain = "C:\share\softpaqs\HPIARepoHead"

# and this for use when selecting a single repository (use with Script version >= 1.15
$RepoShareCommon = "\\$($FileServerName)\share\softpaqs\HPIACommonRepository"
#$RepoShareCommon = "C:\share\softpaqs\HPIACommonRepository"

#-------------------------------------------------------------------
$LogFile = "$PSScriptRoot\HPDriverPackDownload.log"                   # Default to location where script runs from

# if there errors during startup, set this to $true, for additional info
$DebugMode = $false                                                   # Setting managed in GUI (temporarily)

$HPIACMPackage = 'HPIA'             # Package Name for creating/maintaining in SCCM

####################################################################################
#
# Settings managed by Script GUI - these are the defaults read by the script
#
####################################################################################

$OSVER = "1909"                 

$HPIAVersion = '5.0.2.3827'         # info pulled from Browse app in UI

#-------------------------------------------------------------------
# next setting makes the script work with Microsoft SCCM/MECM, if set to $true
# it can be modified in the main script, as needed

# Add ability to create/update HPIA package - chose path from below
$HPIAPath = '\\CM01\Share\Applications\SP107374 HPIA 5.0.2'            
#$HPIAPath = "C:\Share\Applications\HPIA-4.5.8.1"

$UpdateCMPackages = $False
$DistributeCMPackages = $False
$Continueon404 = $True              # 10/20 NEW: setting to keep filters (script version 1.70)
$KeepFilters = $False               # 7/31 NEW: setting to keep filters (script version 1.35)
$CommonRepo = $True                 # 7/21 NEW: manage single repository folder, instead of individual per model
