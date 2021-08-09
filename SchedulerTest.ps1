<#
    reposcriptTest.ps1

    Developed by:
        Dan Felman
        HP Inc Technical Consultant
        9/15/2020

    This is a sample script to show how to schedule the HPIARepo_Downloader script v1.62 and later with runtime options

    'HPIARepo_Downloader' runtime options:

    -IniFile <filepath.ini> -ListFilters -RepoStyle <common|individual> -Sync -NoIniSw -ShowActivityLog -Products <>80D4,8549,8470> -newLog 

        -IniFile   defaults to '.\HPIARepo_ini.ps1'
        -RepoStyle defaults from $CommonRepo variable in IniFile
        -Products  defaults to $HPModels list in IniFile. 
            In either case, corresponding repositories must exist, or product is bypassed

    NOTE: modify location for local environment
          modify options as needed
#>

# this is the PS script block that gets scheduled... modify as needed

$HPIABlock = { 
    & Set-Location 'C:\Users\Administrator.VIAMONSTRA\Documents\Manageability_demo_files\Scripts\WiP'
    & .\HPIARepo_Downloader_1.83.ps1 -ListFilters -RepoStyle individual -ShowActivityLog -Sync -NoIniSw -Products 80D4,8549,8470 -newLog 
}

# let's schedule the Downloader as a Windows Task job
# if job (by name) already exists, just report it, otherwise, schedule it

#--------------------------------------------------------------------------------------

$scheduleTime = '16:56'                           # NOTE: Modify when to schedule job                      

$doJob = $true                                    # default is to schedule the job
$ScheduledJobs = Get-ScheduledJob                 # obtain existing scheduled jobs
$TriggerName = 'HPIATrigger'                      # name of the job... Modify if needed

# find out if the Trigger is already in use

if ( $ScheduledJobs -ne $null ) {
    foreach ( $currJob in $ScheduledJobs ) {
        if ( $currJob.Name -match $TriggerName ) {
            $doJob = $false
            write-host "Trigger found (shown next). Can't schedule another Job with same Trigger name. To remove type: 'Unregister-ScheduledJob -id $($currJob.id)'"
            Get-ScheduledJob -Name $TriggerName
        } # if ( $currJob.Name -match 'HPIATrigger' )
    } # foreach ( $currJob in $jobs ) 
}

# If Trigger not yet scheduled, let's do it now
# but make sure we have admin rights first

if ( $doJob ) {
    $adminRights = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    If( $adminRights ) {
        'Scheduling job. User is a member of local administrators'
        ## NOTE some Options: 
        ##    -Once <-At "1/20/2012 3:00 AM">
        ##    -Weekly -DaysOfWeek <DayOfWeeklist[]> <-At "3:00 AM">
        ##    -Daily <-At "1/20/2012 3:00 AM">
        ##    -AtLogon
        ##    -AtStartup
        ## ( https://docs.microsoft.com/en-us/powershell/module/psscheduledjob/new-jobtrigger?view=powershell-5.1 )
        $trigger = New-JobTrigger -Once -At $scheduleTime          
        Register-ScheduledJob -Name $TriggerName -Trigger $trigger -ScriptBlock $HPIABlock
    } Else {
        'Can not schedule job. User is not a member of local administrators'
    } # else If( $adminRights )
}

#--------------------------------------------------------------------------------------
# Powershell scheduling job commands
#--------------------------------------------------------------------------------------

# Get-ScheduledJob
# Get-ScheduledJob -Name HPIATrigger
# Unregister-ScheduledJob -Name HPIATrigger

# Remove-JobTrigger -Name HPIATrigger