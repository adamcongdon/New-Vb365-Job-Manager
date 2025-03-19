#Requires -Module Veeam.Archiver.PowerShell

[CmdletBinding()]
Param(
    [int]$objectsPerJob = 500,
    [ValidateSet("SharePoint", "Teams", "GroupMailbox")]$limitServiceTo,
    [string]$jobNamePattern = "M365Backup-{0:d3}",
    [switch]$withTeamsChats,
    [switch]$withGroupMailbox,
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
    
    Start-Transcript -Path "${PSScriptRoot}\logs\vb365-m365-jobs-$(Get-Date -Format FileDateTime).log" -NoClobber
}

PROCESS {
    "Starting VB365 Job Manager" | timelog
    "Parameters: limitServiceTo=$limitServiceTo, withGroupMailbox=$withGroupMailbox" | timelog
    
    # Get organization and repositories
    $org = Get-VBOOrganization -Name $PSBoundParameters['Organization']
    $repos = $PSBoundParameters['Repository'] | ForEach-Object { Get-VBORepository -Name $_ }
    
    # Get objects to process
    $sites = if (!$limitServiceTo -or $limitServiceTo -eq "SharePoint") { 
        Get-VBOOrganizationSite -Organization $org -NotInJob 
    } else { @() }
    "Found {0} SharePoint sites" -f $sites.Count | timelog
    
    $teams = if (!$limitServiceTo -or $limitServiceTo -eq "Teams") { 
        Get-VBOOrganizationTeam -NotInJob -Organization $org 
    } else { @() }
    "Found {0} Teams" -f $teams.Count | timelog
    
    $groupMailboxes = if (!$limitServiceTo -or $limitServiceTo -eq "GroupMailbox") { 
        $groups = Get-VBOOrganizationGroup -Organization $org -NotInJob
        "Found {0} GroupMailboxes" -f $groups.Count | timelog
        if ($groups.Count -gt 0) {
            "GroupMailboxes found: {0}" -f ($groups | ForEach-Object { $_.DisplayName + " (SiteUrl: " + $_.SiteUrl + ", Email: " + $_.Email + ")" } | Join-String -Separator ", ") | timelog
        } else {
            "No GroupMailboxes found to process" | timelog
        }
        $groups
    } else { @() }
    
    # Determine primary objects to process
    $objects = switch ($limitServiceTo) {
        "Teams" { $teams }
        "GroupMailbox" { $groupMailboxes }
        default { $sites }
    }
    "Processing {0} primary objects" -f $objects.Count | timelog
    
    $jobNum = 1
    $currentJob = $null
    $currentSchedule = $baseSchedule
    $repoIndex = 0
    $objCount = 0
    
    foreach ($obj in $objects) {
        "Processing: {0} (URL: {1})" -f $obj.ToString(), $obj.URL | timelog
        
        # Apply filters
        if ($includes -and !($includes | Where-Object { $obj.toString() -cmatch $_ })) { 
            "Skipping {0} - no include match" -f $obj.ToString() | timelog
            continue 
        }
        if ($excludes | Where-Object { $obj.toString() -cmatch $_ }) { 
            "Skipping {0} - exclude match" -f $obj.ToString() | timelog
            continue 
        }
        
        # Create or get job with initial item
        if (!$currentJob -or $objCount -ge $objectsPerJob) {
            $jobName = $jobNamePattern -f $jobNum++
            $repo = $repos[$repoIndex++ % $repos.Count]
            $currentJob = Get-VBOJob -Name $jobName -Organization $org
            
            if (!$currentJob) {
                # Create initial backup item based on object type
                $initialItem = switch ($limitServiceTo) {
                    "Teams" { New-VBOBackupItem -Team $obj -TeamsChats:$withTeamsChats }
                    "GroupMailbox" { New-VBOBackupItem -Group $obj -GroupMailbox:$withGroupMailbox }
                    default { New-VBOBackupItem -Site $obj }
                }
                "Creating new job {0} with initial item: {1}" -f $jobName, $initialItem.ToString() | timelog
                
                $currentJob = Add-VBOJob -Organization $org -Name $jobName -Repository $repo `
                    -SchedulePolicy $currentSchedule -SelectedItems $initialItem
                
                $currentSchedule = New-VBOJobSchedulePolicy -EnableSchedule -Type $currentSchedule.Type `
                    -DailyType $currentSchedule.DailyType -DailyTime ($currentSchedule.DailyTime + $scheduleDelay)
                $objCount = 1
            } else {
                $objCount = (Get-VBOBackupItem -Job $currentJob).Count
                "Using existing job {0} with {1} items" -f $jobName, $objCount | timelog
            }
        }
        
        # Add items to the job (whether new or existing)
        switch ($limitServiceTo) {
            "Teams" {
                $item = New-VBOBackupItem -Team $obj -TeamsChats:$withTeamsChats
                try {
                    Add-VBOBackupItem -Job $currentJob -BackupItem $item -ErrorAction Stop
                    "Added Team: {0}" -f $obj.ToString() | timelog
                    $objCount++
                } catch {
                    "Failed to add Team {0}: {1}" -f $obj.ToString(), $_.Exception.Message | timelog
                }
            }
            "GroupMailbox" {
                $item = New-VBOBackupItem -Group $obj -GroupMailbox:$withGroupMailbox
                try {
                    Add-VBOBackupItem -Job $currentJob -BackupItem $item -ErrorAction Stop
                    "Added GroupMailbox: {0} (Mailbox included: {1})" -f $obj.ToString(), $withGroupMailbox | timelog
                    $objCount++
                } catch {
                    "Failed to add GroupMailbox {0}: {1}" -f $obj.ToString(), $_.Exception.Message | timelog
                }
            }
            default { # SharePoint or no limit
                $siteItem = New-VBOBackupItem -Site $obj
                try {
                    Add-VBOBackupItem -Job $currentJob -BackupItem $siteItem -ErrorAction Stop
                    "Added Site: {0}" -f $obj.ToString() | timelog
                    $objCount++
                } catch {
                    "Failed to add Site {0}: {1}" -f $obj.ToString(), $_.Exception.Message | timelog
                }
                
                if (!$limitServiceTo) {
                    # Add matching Team if exists
                    $team = $teams | Where-Object { ($_.Mail -split "@")[0] -eq ([uri]$obj.URL).Segments[-1] }
                    if ($team) {
                        $teamItem = New-VBOBackupItem -Team $team -TeamsChats:$withTeamsChats
                        try {
                            Add-VBOBackupItem -Job $currentJob -BackupItem $teamItem -ErrorAction Stop
                            "Added matching Team: {0}" -f $team.ToString() | timelog
                            $objCount++
                        } catch {
                            "Failed to add Team {0}: {1}" -f $team.ToString(), $_.Exception.Message | timelog
                        }
                    }
                    # Add matching GroupMailbox if exists
                    $siteUrl = ([uri]$obj.URL).AbsolutePath.TrimEnd('/')
                    $group = $null
                    $group = $groupMailboxes | Where-Object { 
                        $groupSiteUrl = if ($_.SiteUrl) { ([uri]$_.SiteUrl).AbsolutePath.TrimEnd('/') } else { $null }
                        $groupSiteUrl -eq $siteUrl
                    } | Select-Object -First 1
                    if (-not $group) {
                        $siteName = $obj.ToString().Replace(" ", "").ToLower()
                        $group = $groupMailboxes | Where-Object {
                            if ($_.DisplayName) {
                                $groupName = $_.DisplayName.Replace(" ", "").Replace("M365Group", "").ToLower()
                                $groupName -eq $siteName -or $groupName -like "*$siteName*"
                            } else {
                                $false
                            }
                        } | Select-Object -First 1
                    }
                    if ($group -and $group -is [Veeam.Archiver.PowerShell.Model.VBOOrganizationGroup]) {
                        $groupItem = New-VBOBackupItem -Group $group -GroupMailbox:$withGroupMailbox
                        try {
                            Add-VBOBackupItem -Job $currentJob -BackupItem $groupItem -ErrorAction Stop
                            "Added matching GroupMailbox: {0} (Mailbox included: {1})" -f $group.Email, $withGroupMailbox | timelog
                            $objCount++
                        } catch {
                            "Failed to add GroupMailbox {0}: {1}" -f $group.Email, $_.Exception.Message | timelog
                        }
                    } else {
                        "No matching GroupMailbox found for site {0} (Normalized URL: {1}, Name: {2})" -f $obj.ToString(), $siteUrl, $siteName | timelog
                        if ($groupMailboxes.Count -gt 0) {
                            "Available GroupMailboxes: {0}" -f ($groupMailboxes | ForEach-Object { 
                                $normalizedName = if ($_.DisplayName) { $_.DisplayName.Replace(" ", "").Replace("M365Group", "").ToLower() } else { "N/A" }
                                "$($_.DisplayName) (Normalized Name: $normalizedName, SiteUrl: $($_.SiteUrl), Email: $($_.Email))"
                            } | Join-String -Separator ", ") | timelog
                        }
                    }
                }
            }
        }
    }
}

END {
    Stop-Transcript
}