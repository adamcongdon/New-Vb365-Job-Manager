#Requires -Module Veeam.Archiver.PowerShell

[CmdletBinding()]
Param(
    [int]$objectsPerJob = 500,
    [ValidateSet("SharePoint", "Teams")]$limitServiceTo,
    [string]$jobNamePattern = "SharePointTeams-{0:d3}",
    [switch]$withTeamsChats,
    [Object]$baseSchedule,
    [string]$scheduleDelay = "00:30:00",
    [string]$includeFile,
    [string]$excludeFile,
    [switch]$recurseSP,
    [switch]$checkBackups,
    [int]$countTeamAs = 3
)

DynamicParam {
    Import-Module Veeam.Archiver.PowerShell
    $dict = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
    
    # Organization parameter
    $orgAttr = New-Object System.Management.Automation.ParameterAttribute
    $orgAttr.Mandatory = $true
    $orgSet = Get-VBOOrganization | Select-Object -ExpandProperty Name
    $orgValid = New-Object System.Management.Automation.ValidateSetAttribute($orgSet)
    $dict.Add('Organization', (New-Object System.Management.Automation.RuntimeDefinedParameter('Organization', [string], @($orgAttr, $orgValid))))
    
    # Repository parameter
    $repoAttr = New-Object System.Management.Automation.ParameterAttribute
    $repoAttr.Mandatory = $true
    $repoSet = Get-VBORepository | Select-Object -ExpandProperty Name
    $repoValid = New-Object System.Management.Automation.ValidateSetAttribute($repoSet)
    $dict.Add('Repository', (New-Object System.Management.Automation.RuntimeDefinedParameter('Repository', [string[]], @($repoAttr, $repoValid))))
    
    return $dict
}

BEGIN {
    # Basic logging
    filter timelog { "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss"): $_" }
    
    # Set defaults
    if (!$baseSchedule) {
        $baseSchedule = New-VBOJobSchedulePolicy -EnableSchedule -Type Daily -DailyType Everyday -DailyTime "22:00:00"
    }
    
    # Load include/exclude files
    $basename = $MyInvocation.MyCommand.Name.Split(".")[0]
    $includeFile = if ($includeFile) { $includeFile } elseif (Test-Path "${PSScriptRoot}\${basename}.includes") { "${PSScriptRoot}\${basename}.includes" }
    $excludeFile = if ($excludeFile) { $excludeFile } elseif (Test-Path "${PSScriptRoot}\${basename}.excludes") { "${PSScriptRoot}\${basename}.excludes" }
    $includes = if ($includeFile) { Get-Content $includeFile } else { @() }
    $excludes = if ($excludeFile) { Get-Content $excludeFile } else { @() }
    
    Start-Transcript -Path "${PSScriptRoot}\logs\vb365-spo-teams-jobs-$(Get-Date -Format FileDateTime).log" -NoClobber
}

PROCESS {
    "Starting VB365 Job Manager" | timelog
    
    # Get organization and repositories
    $org = Get-VBOOrganization -Name $PSBoundParameters['Organization']
    $repos = $PSBoundParameters['Repository'] | ForEach-Object { Get-VBORepository -Name $_ }
    
    # Get objects to process
    $sites = if (!$limitServiceTo -or $limitServiceTo -eq "SharePoint") { 
        Get-VBOOrganizationSite -Organization $org -NotInJob 
    } else { @() }
    
    $teams = if (!$limitServiceTo -or $limitServiceTo -eq "Teams") { 
        Get-VBOOrganizationTeam -NotInJob -Organization $org 
    } else { @() }
    
    # Process objects
    $objects = if ($limitServiceTo -eq "Teams") { $teams } else { $sites }
    $jobNum = 1
    $currentJob = $null
    $currentSchedule = $baseSchedule
    $repoIndex = 0
    $objCount = 0
    
    foreach ($obj in $objects) {
        # Apply filters
        if ($includes -and !($includes | Where-Object { $obj.toString() -cmatch $_ })) { continue }
        if ($excludes | Where-Object { $obj.toString() -cmatch $_ }) { continue }
        
        # Create or get job
        if (!$currentJob -or $objCount -ge $objectsPerJob) {
            $jobName = $jobNamePattern -f $jobNum++
            $repo = $repos[$repoIndex++ % $repos.Count]
            $currentJob = Get-VBOJob -Name $jobName -Organization $org
            if (!$currentJob) {
                $currentJob = Add-VBOJob -Organization $org -Name $jobName -Repository $repo -SchedulePolicy $currentSchedule
                $currentSchedule = New-VBOJobSchedulePolicy -EnableSchedule -Type $currentSchedule.Type `
                    -DailyType $currentSchedule.DailyType -DailyTime ($currentSchedule.DailyTime + $scheduleDelay)
            }
            $objCount = 0
        }
        
        # Add object to job
        if ($limitServiceTo -eq "Teams") {
            Add-VBOBackupItem -Job $currentJob -BackupItem (New-VBOBackupItem -Team $obj -TeamsChats:$withTeamsChats)
        } else {
            Add-VBOBackupItem -Job $currentJob -BackupItem (New-VBOBackupItem -Site $obj)
            if (!$limitServiceTo) {
                $team = $teams | Where-Object { ($_.Mail -split "@")[0] -eq ([uri]$obj.URL).Segments[-1] }
                if ($team) {
                    Add-VBOBackupItem -Job $currentJob -BackupItem (New-VBOBackupItem -Team $team -TeamsChats:$withTeamsChats)
                }
            }
        }
        $objCount++
    }
}

END {
    Stop-Transcript
}