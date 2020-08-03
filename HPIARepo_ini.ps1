<#
    File: HPIARepo_ini.ps1

    Dan Felman/HP Inc
    7/27/2020

    modify variables as need for use by the Downloader script
#>

$OS = "Win10"
$OSVER = "1909"
$OSVALID = @("1809", "1903", "1909", "2004")

$FilterCategories = @('Driver','BIOS', 'Firmware', 'Software')

#-------------------------------------------------------------------
# Example systems to manage
#
# NEW 1.30 - added list of individual softpaq names to host in repository
$NBSet1 = 'HP Collaboration Keyboard Software' , 'HP Hotkey Support' , 'System Default Settings for Windows 10' , 'HP Notifications'
$NBSet2 = 'HP Hotkey Support' , 'System Default Settings for Windows 10' , 'HP Notifications'
$DTSet1 = 'HP Notifications'

$HPModelsTable = @(
	@{ ProdCode = '8598'; Model = "HP ProDesk 600 G5 DM";                   SqName = $DTSet1  }
	@{ ProdCode = '8438'; Model = "HP EliteBook x360 1030 G3";              SqName = $NBSet1  }
	@{ ProdCode = '83B3'; Model = "HP ELITEBOOK 830 G5";                    SqName = $NBSet1  }
	@{ ProdCode = '83B2'; Model = "HP ELITEBOOK 840 G5";                    SqName = $NBSet1  }
	@{ ProdCode = '8549'; Model = "HP ELITEBOOK 840 G6";                    SqName = $NBSet2  }
	@{ ProdCode = '8470'; Model = "HP ELITEBOOK X360 1040 G5";              SqName = $NBSet1  }
	@{ ProdCode = '844F'; Model = "HP ZBook Studio x360 G5";                SqName = $NBSet1  }
	@{ ProdCode = '8593'; Model = "HP EliteDesk 800 G5 Desktop Mini PC";    SqName = $DTSet1  }
	@{ ProdCode = '8549'; Model = "HP EliteBook 840 G6 Healthcare Edition"; SqName = $DTSet1  }
	@{ ProdCode = '859F'; Model = "HP EliteOne 800 G5 All-in-One";          SqName = $DTSet1  }
    )

$FileServerName = $Env:COMPUTERNAME 

# 7/28 NEW: Version 1.31 - Add ability to create/update HPIA package - chose path from below
#$HPIAPath = "\\$($FileServerName)\Share\Applications\HPIA-4.5.8.1"
$HPIAPath = "C:\Share\Applications\HPIA-4.5.8.1"

$HPIAPackage = 'HPIA'
$HPIAVersion = '4.5.8.1'

#-------------------------------------------------------------------
# choose/edit a path from next 2 lines for the location of the main repository folder

# next for use in multiple repositories/one per model
#$RepoShareMain = $RepoShareMain = "\\$($FileServerName)\share\softpaqs\HPIARepoHead"
$RepoShareMain = "C:\share\softpaqs\HPIARepoHead"


# 7/23 NEW: and this for use when selecting a single repository (use with Script version >= 1.15
#$RepoShareCommon = "\\$($FileServerName)\share\softpaqs\HPIACommonRepository"
$RepoShareCommon = "C:\share\softpaqs\HPIACommonRepository"

#-------------------------------------------------------------------
$LogFile = "$PSScriptRoot\HPDriverPackDownload.log"               # Default to location where script runs from

#-------------------------------------------------------------------
# next setting makes the script work with Microsoft SCCM/MECM, if set to $true
# it can be modified in the main script, as needed
$UpdateCMPackages = $False
$DistributeCMPackages = $False

#-------------------------------------------------------------------
# 7/31 NEW: setting to keep filters (script version 1.35)
$KeepFilters = $False

#-------------------------------------------------------------------
# 7/21 NEW: manage single repository folder, instead of individual per model
$CommonRepo = $False

# if there errors during startup, set this to $true, for additional info
$DebugMode = $false
