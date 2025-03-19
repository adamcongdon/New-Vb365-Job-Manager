#Requires -Module Veeam.Archiver.PowerShell

[CmdletBinding()]
Param(
    [Parameter()]
    [ValidateScript({
        if ($_ % 3 -ne 0) { throw "objectsPerJob must be a multiple of 3" }
        if ($_ -lt 3 -or $_ -gt 3000) { throw "objectsPerJob must be between 3 and 3000" }
        $true
    })]
    [int]$objectsPerJob = 600,
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
    filter timelog { 
        $message = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss"): $_"
        Write-Host $message # Write to host/console, which gets captured by Start-Transcript
    }
    
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
    $jobAssignments = @{}
    
    "Starting VB365 Job Manager" | timelog
    "Parameters: limitServiceTo=$limitServiceTo, withGroupMailbox=$withGroupMailbox, objectsPerJob=$objectsPerJob" | timelog
    
    # Get organization and repositories
    $org = Get-VBOOrganization -Name $PSBoundParameters['Organization']
    $repos = $PSBoundParameters['Repository'] | ForEach-Object { Get-VBORepository -Name $_ }
    
    # Get objects to process
    $sites = if (!$limitServiceTo -or $limitServiceTo -eq "SharePoint") { 
        $rawSites = Get-VBOOrganizationSite -Organization $org -NotInJob
        $validSites = $rawSites | Where-Object { 
            $_ -and $_ -is [Veeam.Archiver.PowerShell.Model.VBOOrganizationSite] -and $_.ToString() -and $_.URL
        }
        $invalidSites = $rawSites | Where-Object { 
            -not ($_ -and $_ -is [Veeam.Archiver.PowerShell.Model.VBOOrganizationSite] -and $_.ToString() -and $_.URL)
        }
        if ($invalidSites) {
            "Debug: Found {0} invalid SharePoint sites: {1}" -f $invalidSites.Count, ($invalidSites | ForEach-Object { "Site: $_" } | Join-String -Separator ", ") | timelog
        }
        "Found {0} SharePoint sites not in jobs" -f $validSites.Count | timelog
        if ($validSites.Count -gt 0) {
            "SharePoint sites not in jobs: {0}" -f ($validSites | ForEach-Object { 
                "$($_.ToString()) (URL: $($_.URL))"
            } | Join-String -Separator ", ") | timelog
        }
        $validSites
    } else { @() }

    $teams = if (!$limitServiceTo -or $limitServiceTo -eq "Teams") { 
        $rawTeams = Get-VBOOrganizationTeam -NotInJob -Organization $org
        $validTeams = $rawTeams | Where-Object { 
            $_ -and $_ -is [Veeam.Archiver.PowerShell.Model.VBOOrganizationTeam] -and $_.ToString() -and $_.Mail
        }
        $invalidTeams = $rawTeams | Where-Object { 
            -not ($_ -and $_ -is [Veeam.Archiver.PowerShell.Model.VBOOrganizationTeam] -and $_.ToString() -and $_.Mail)
        }
        if ($invalidTeams) {
            "Debug: Found {0} invalid Teams: {1}" -f $invalidTeams.Count, ($invalidTeams | ForEach-Object { "Team: $_" } | Join-String -Separator ", ") | timelog
        }
        "Found {0} Teams not in jobs" -f $validTeams.Count | timelog
        if ($validTeams.Count -gt 0) {
            "Teams not in jobs: {0}" -f ($validTeams | ForEach-Object { 
                "$($_.ToString()) (Email: $($_.Mail))"
            } | Join-String -Separator ", ") | timelog
        }
        $validTeams
    } else { @() }

    $groupMailboxes = if (!$limitServiceTo -or $limitServiceTo -eq "GroupMailbox") { 
        $rawGroups = Get-VBOOrganizationGroup -Organization $org -NotInJob
        $validGroups = $rawGroups | Where-Object { 
            $_ -and $_ -is [Veeam.Archiver.PowerShell.Model.VBOOrganizationGroup] -and $_.ToString() -and $_.GroupName -and $_.GroupName -ne '' -and $_.DisplayName -and $_.DisplayName -ne ''
        }
        $invalidGroups = $rawGroups | Where-Object { 
            -not ($_ -and $_ -is [Veeam.Archiver.PowerShell.Model.VBOOrganizationGroup] -and $_.ToString() -and $_.GroupName -and $_.GroupName -ne '' -and $_.DisplayName -and $_.DisplayName -ne '')
        }
        if ($invalidGroups) {
            "Debug: Found {0} invalid GroupMailboxes: {1}" -f $invalidGroups.Count, ($invalidGroups | ForEach-Object { 
                "Group: $_, Properties: $($_.PSObject.Properties | ForEach-Object { $_.Name + '=' + $_.Value } | Join-String -Separator ', ')"
            } | Join-String -Separator "; ") | timelog
        }
        "Found {0} GroupMailboxes not in jobs" -f $validGroups.Count | timelog
        if ($validGroups.Count -gt 0) {
            "GroupMailboxes not in jobs: {0}" -f ($validGroups | ForEach-Object { 
                $name = if ($_.DisplayName) { $_.DisplayName } else { "Unnamed" }
                "Debug: GroupMailbox - Name: $name, GroupName: $($_.GroupName), Properties: $($_.PSObject.Properties | ForEach-Object { $_.Name + '=' + $_.Value } | Join-String -Separator ', ')" | timelog
                "$name (SiteUrl: $($_.SiteUrl), GroupName: $($_.GroupName))"
            } | Join-String -Separator ", ") | timelog
        } else {
            "No GroupMailboxes found to process" | timelog
        }
        $validGroups
    } else { @() }

    # Check if there are no items to process
    if ($sites.Count -eq 0 -and $teams.Count -eq 0 -and $groupMailboxes.Count -eq 0) {
        "No items (SharePoint sites, Teams, or GroupMailboxes) found to process. All items may already be in jobs." | timelog
        return
    }

    # Determine primary objects to process
    $objects = switch ($limitServiceTo) {
        "Teams" { 
            $teams | ForEach-Object { 
                if ($_.Mail -and $_.DisplayName) {
                    [PSCustomObject]@{ 
                        Type = "Team"; 
                        Object = [PSCustomObject]@{ 
                            Name = $_.DisplayName; 
                            Identifier = $_.Mail 
                        } 
                    }
                } else {
                    "Debug: Skipping invalid Team object: $_" | timelog
                    $null
                }
            } | Where-Object { $_ }
        }
        "GroupMailbox" { 
            $groupMailboxes | ForEach-Object { 
                if ($_.GroupName -and $_.DisplayName) {
                    [PSCustomObject]@{ 
                        Type = "GroupMailbox"; 
                        Object = [PSCustomObject]@{ 
                            Name = $_.DisplayName; 
                            Identifier = $_.GroupName 
                        } 
                    }
                } else {
                    "Debug: Skipping invalid GroupMailbox object: $_" | timelog
                    $null
                }
            } | Where-Object { $_ }
        }
        default { 
            if (!$limitServiceTo) {
                $combined = @()
                $combined += $sites | ForEach-Object { 
                    if ($_.URL -and $_.Name) {
                        [PSCustomObject]@{ 
                            Type = "Site"; 
                            OriginalObject = $_;  # Store the original for later use
                            Object = [PSCustomObject]@{ 
                                Name = $_.Name; 
                                Identifier = $_.URL 
                            } 
                        }
                    } else {
                        "Debug: Skipping invalid Site object: $_" | timelog
                        $null
                    }
                }
                $combined += $teams | ForEach-Object { 
                    if ($_.Mail -and $_.DisplayName) {
                        [PSCustomObject]@{ 
                            Type = "Team"; 
                            OriginalObject = $_; 
                            Object = [PSCustomObject]@{ 
                                Name = $_.DisplayName; 
                                Identifier = $_.Mail 
                            } 
                        }
                    } else {
                        "Debug: Skipping invalid Team object: $_" | timelog
                        $null
                    }
                }
                $combined += $groupMailboxes | ForEach-Object { 
                    if ($_.GroupName -and $_.DisplayName) {
                        [PSCustomObject]@{ 
                            Type = "GroupMailbox"; 
                            OriginalObject = $_; 
                            Object = [PSCustomObject]@{ 
                                Name = $_.DisplayName; 
                                Identifier = $_.GroupName 
                            } 
                        }
                    } else {
                        "Debug: Skipping invalid GroupMailbox object: $_" | timelog
                        $null
                    }
                } | Where-Object { $_ }
                "Debug: Raw objects count: $($combined.Count)" | timelog
                $combined | ForEach-Object { "Debug: Object Type: $($_.Type), Object: $($_.Object.Name) (Identifier: $($_.Object.Identifier))" | timelog }
                $combined
            } else {
                $sites | ForEach-Object { 
                    if ($_.URL -and $_.Name) {
                        [PSCustomObject]@{ 
                            Type = "Site"; 
                            Object = [PSCustomObject]@{ 
                                Name = $_.Name; 
                                Identifier = $_.URL 
                            } 
                        }
                    } else {
                        "Debug: Skipping invalid Site object: $_" | timelog
                        $null
                    }
                } | Where-Object { $_ }
            }
        }
    }
    "Processing {0} primary objects" -f $objects.Count | timelog
    if ($objects.Count -gt 0) {
        "Primary objects to process: {0}" -f ($objects | ForEach-Object {
    if (-not $_.Object -or -not ($_.Type -in @("Site", "Team", "GroupMailbox"))) {
        "Skipping invalid object of type $($_.Type) with Object: $($_.Object)"
    } else {
        switch ($_.Type) {
            "Site" { "Site: $($_.Object.Name) (URL: $($_.Object.Identifier))" }
            "Team" { "Team: $($_.Object.Name) (Email: $($_.Object.Identifier))" }
            "GroupMailbox" { "GroupMailbox: $($_.Object.Name) (GroupName: $($_.Object.Identifier))" }
        }
    }
} | Where-Object { $_ } | Join-String -Separator ", ") | timelog
    }
    
    # Get all existing jobs matching the pattern
    $existingJobs = @()
    $jobNum = 1
    while ($true) {
        $jobName = $jobNamePattern -f $jobNum
        $job = Get-VBOJob -Name $jobName -Organization $org
        if ($job) {
            $existingJobs += $job
            $jobNum++
        } else {
            break
        }
    }
    "Found {0} existing jobs" -f $existingJobs.Count | timelog
    
    # Check existing jobs for missing GroupMailboxes
    foreach ($job in $existingJobs) {
        "Checking existing job {0} for missing GroupMailboxes" -f $job.Name | timelog
        $jobItems = Get-VBOBackupItem -Job $job
        $jobSites = $jobItems | Where-Object { $_.Type -eq "Site" }
        $jobTeams = $jobItems | Where-Object { $_.Type -eq "Team" }
        $jobGroups = $jobItems | Where-Object { $_.Type -eq "Group" }
        
        # Check Sites for missing GroupMailboxes
        foreach ($site in $jobSites) {
            $siteObj = $site.Site
            $siteUrl = ([uri]$siteObj.URL).AbsolutePath.TrimEnd('/')
            $siteName = $siteObj.ToString()
            
            # Check if there's a matching GroupMailbox in the job
            $hasGroup = $jobGroups | Where-Object { 
                $groupSiteUrl = if ($_.Group.SiteUrl) { ([uri]$_.Group.SiteUrl).AbsolutePath.TrimEnd('/') } else { $null }
                $groupSiteUrl -eq $siteUrl -or ($_.Group.DisplayName -and $_.Group.DisplayName.Replace(" ", "").Replace("M365Group", "").ToLower() -eq $siteName.Replace(" ", "").ToLower())
            }
            
            if (-not $hasGroup) {
                # Find the matching GroupMailbox from the available list
                $group = $groupMailboxes | Where-Object { 
                    $groupSiteUrl = if ($_.SiteUrl) { ([uri]$_.SiteUrl).AbsolutePath.TrimEnd('/') } else { $null }
                    $groupSiteUrl -eq $siteUrl
                } | Select-Object -First 1
                if (-not $group) {
                    $group = $groupMailboxes | Where-Object {
                        if ($_.DisplayName) {
                            $groupName = $_.DisplayName.Replace(" ", "").Replace("M365Group", "").ToLower()
                            $siteNormalized = $siteName.Replace(" ", "").ToLower()
                            $groupName -eq $siteNormalized -or $groupName -like "*$siteNormalized*"
                        } else {
                            $false
                        }
                    } | Select-Object -First 1
                }
                
                if ($group -and $group -is [Veeam.Archiver.PowerShell.Model.VBOOrganizationGroup]) {
                    $groupItem = New-VBOBackupItem -Group $group -GroupMailbox:$withGroupMailbox
                    try {
                        $groupIdentifier = if ($group.DisplayName) { $group.DisplayName } else { $group.GroupName }
                        Add-VBOBackupItem -Job $job -BackupItem $groupItem -ErrorAction Stop
                        "Added missing GroupMailbox {0} (GroupName: {1}) to job {2} (Mailbox included: {3})" -f $groupIdentifier, $group.GroupName, $job.Name, $withGroupMailbox | timelog
                        if (-not $jobAssignments[$job.Name]) { $jobAssignments[$job.Name] = @() }
                        $jobAssignments[$job.Name] += "GroupMailbox: $groupIdentifier (GroupName: $($group.GroupName))"
                        # Remove from $groupMailboxes to avoid re-adding later
                        $groupMailboxes = $groupMailboxes | Where-Object { $_.GroupName -ne $group.GroupName }
                    } catch {
                        "Failed to add missing GroupMailbox {0} (GroupName: {1}) to job {2}: {3}" -f $groupIdentifier, $group.GroupName, $job.Name, $_.Exception.Message | timelog
                    }
                }
            }
        }
        
        # Check Teams for missing GroupMailboxes
        foreach ($team in $jobTeams) {
            $teamName = $team.Team.ToString()
            $teamEmail = ($team.Team.Mail -split "@")[0]
            
            # Check if there's a matching GroupMailbox in the job
            $hasGroup = $jobGroups | Where-Object { 
                $groupName = if ($_.Group.DisplayName) { $_.Group.DisplayName.Replace(" ", "").Replace("M365Group", "").ToLower() } else { $null }
                $groupName -eq $teamName.Replace(" ", "").ToLower()
            }
            
            if (-not $hasGroup) {
                $group = $groupMailboxes | Where-Object {
                    if ($_.DisplayName) {
                        $groupName = $_.DisplayName.Replace(" ", "").Replace("M365Group", "").ToLower()
                        $teamNormalized = $teamName.Replace(" ", "").ToLower()
                        $groupName -eq $teamNormalized -or $groupName -like "*$teamNormalized*"
                    } else {
                        $false
                    }
                } | Select-Object -First 1
                
                if ($group -and $group -is [Veeam.Archiver.PowerShell.Model.VBOOrganizationGroup]) {
                    $groupItem = New-VBOBackupItem -Group $group -GroupMailbox:$withGroupMailbox
                    try {
                        $groupIdentifier = if ($group.DisplayName) { $group.DisplayName } else { $group.GroupName }
                        Add-VBOBackupItem -Job $job -BackupItem $groupItem -ErrorAction Stop
                        "Added missing GroupMailbox {0} (GroupName: {1}) to job {2} (Mailbox included: {3})" -f $groupIdentifier, $group.GroupName, $job.Name, $withGroupMailbox | timelog
                        if (-not $jobAssignments[$job.Name]) { $jobAssignments[$job.Name] = @() }
                        $jobAssignments[$job.Name] += "GroupMailbox: $groupIdentifier (GroupName: $($group.GroupName))"
                        # Remove from $groupMailboxes to avoid re-adding later
                        $groupMailboxes = $groupMailboxes | Where-Object { $_.GroupName -ne $group.GroupName }
                    } catch {
                        "Failed to add missing GroupMailbox {0} (GroupName: {1}) to job {2}: {3}" -f $groupIdentifier, $group.GroupName, $job.Name, $_.Exception.Message | timelog
                    }
                }
            }
        }
    }
    
    # Process new items
    $currentJob = $null
    $currentSchedule = $baseSchedule
    $repoIndex = 0
    $objCount = 0
    
    foreach ($obj in $objects) {
        if (-not $obj.OriginalObject -or -not ($obj.Type -in @("Site", "Team", "GroupMailbox"))) {
            "Skipping invalid object in processing list: $($obj.Type) - $($obj.OriginalObject)" | timelog
            continue
        }
        
        $itemType = $obj.Type
        $itemObject = $obj.OriginalObject
        
        switch ($itemType) {
            "Site" { "Processing Site: {0} (URL: {1})" -f $itemObject.ToString(), $itemObject.URL | timelog }
            "Team" { "Processing Team: {0} (Email: {1})" -f $itemObject.ToString(), $itemObject.Mail | timelog }
            "GroupMailbox" { 
                $name = if ($itemObject.DisplayName) { $itemObject.DisplayName } else { "Unnamed" }
                "Processing GroupMailbox: {0} (GroupName: {1})" -f $name, $itemObject.GroupName | timelog
            }
        }
        
        # Apply filters
        if ($includes -and !($includes | Where-Object { $itemObject.toString() -cmatch $_ })) { 
            "Skipping {0} - no include match" -f $itemObject.ToString() | timelog
            continue 
        }
        if ($excludes | Where-Object { $itemObject.toString() -cmatch $_ }) { 
            "Skipping {0} - exclude match" -f $itemObject.ToString() | timelog
            continue 
        }
        
        # Find a job that already contains a related item (if any)
        $relatedJob = $null
        if ($itemType -eq "Team") {
            $teamName = ($itemObject.Mail -split "@")[0].Replace(" ", "").ToLower()
            $relatedJob = $existingJobs | Where-Object {
                $jobItems = Get-VBOBackupItem -Job $_
                $jobSites = $jobItems | Where-Object { $_.Type -eq "Site" }
                $jobGroups = $jobItems | Where-Object { $_.Type -eq "Group" }
                ($jobSites | Where-Object { ([uri]$_.Site.URL).Segments[-1].Replace(" ", "").ToLower() -eq $teamName }) -or
                ($jobGroups | Where-Object { $_.Group.DisplayName.Replace(" ", "").Replace("M365Group", "").ToLower() -eq $teamName })
            } | Select-Object -First 1
        } elseif ($itemType -eq "GroupMailbox") {
            if ($itemObject -and $itemObject.DisplayName) {
                $groupName = $itemObject.DisplayName.Replace(" ", "").Replace("M365Group", "").ToLower()
                $relatedJob = $existingJobs | Where-Object {
                    $jobItems = Get-VBOBackupItem -Job $_
                    $jobSites = $jobItems | Where-Object { $_.Type -eq "Site" }
                    $jobTeams = $jobItems | Where-Object { $_.Type -eq "Team" }
                    ($jobSites | Where-Object { $_.Site.ToString().Replace(" ", "").ToLower() -eq $groupName }) -or
                    ($jobTeams | Where-Object { ($_.Team.Mail -split "@")[0].Replace(" ", "").ToLower() -eq $groupName })
                } | Select-Object -First 1
            }
        } else { # SharePoint
            $siteUrl = ([uri]$itemObject.URL).AbsolutePath.TrimEnd('/')
            $siteName = $itemObject.ToString().Replace(" ", "").ToLower()
            $relatedJob = $existingJobs | Where-Object {
                $jobItems = Get-VBOBackupItem -Job $_
                $jobTeams = $jobItems | Where-Object { $_.Type -eq "Team" }
                $jobGroups = $jobItems | Where-Object { $_.Type -eq "Group" }
                ($jobTeams | Where-Object { ($_.Team.Mail -split "@")[0].Replace(" ", "").ToLower() -eq $siteName }) -or
                ($jobGroups | Where-Object { $_.Group.DisplayName.Replace(" ", "").Replace("M365Group", "").ToLower() -eq $siteName })
            } | Select-Object -First 1
        }
        
        if ($relatedJob) {
            $currentJob = $relatedJob
            $objCount = (Get-VBOBackupItem -Job $currentJob).Count
            "Using related job {0} with {1} items" -f $currentJob.Name, $objCount | timelog
        } elseif (!$currentJob -or $objCount -ge $objectsPerJob) {
            $jobName = $jobNamePattern -f $jobNum++
            $repo = $repos[$repoIndex++ % $repos.Count]
            $currentJob = Get-VBOJob -Name $jobName -Organization $org
            
            if (!$currentJob) {
                $initialItem = $null
                switch ($itemType) {
                    "Team" { $initialItem = New-VBOBackupItem -Team $itemObject -TeamsChats:$withTeamsChats }
                    "GroupMailbox" { $initialItem = New-VBOBackupItem -Group $itemObject -GroupMailbox:$withGroupMailbox }
                    default { $initialItem = New-VBOBackupItem -Site $itemObject }
                }
                if ($initialItem) {
                    switch ($itemType) {
                        "Site" { "Creating new job {0} with initial item: Site {1} (URL: {2})" -f $jobName, $itemObject.ToString(), $itemObject.URL | timelog }
                        "Team" { "Creating new job {0} with initial item: Team {1} (Email: {2})" -f $jobName, $itemObject.ToString(), $itemObject.Mail | timelog }
                        "GroupMailbox" { 
                            $name = if ($itemObject.DisplayName) { $itemObject.DisplayName } else { "Unnamed" }
                            "Creating new job {0} with initial item: GroupMailbox {1} (GroupName: {2})" -f $jobName, $name, $itemObject.GroupName | timelog
                        }
                    }
                    $currentJob = Add-VBOJob -Organization $org -Name $jobName -Repository $repo `
                        -SchedulePolicy $currentSchedule -SelectedItems $initialItem
                    $currentSchedule = New-VBOJobSchedulePolicy -EnableSchedule -Type $currentSchedule.Type `
                        -DailyType $currentSchedule.DailyType -DailyTime ($currentSchedule.DailyTime + $scheduleDelay)
                    $objCount = 1
                    $existingJobs += $currentJob
                } else {
                    "Failed to create job {0}: Initial item is null" -f $jobName | timelog
                    continue
                }
            } else {
                $objCount = (Get-VBOBackupItem -Job $currentJob).Count
                "Using existing job {0} with {1} items" -f $jobName, $objCount | timelog
            }
        }
        
        if (-not $currentJob) {
            "Skipping item {0} due to failure in job creation" -f $itemObject.ToString() | timelog
            continue
        }
        
        # Add items to the job
        switch ($itemType) {
            "Team" {
                $item = New-VBOBackupItem -Team $itemObject -TeamsChats:$withTeamsChats
                try {
                    Add-VBOBackupItem -Job $currentJob -BackupItem $item -ErrorAction Stop
                    "Added Team {0} (Email: {1}) to job {2}" -f $itemObject.ToString(), $itemObject.Mail, $currentJob.Name | timelog
                    $objCount++
                    if (-not $jobAssignments[$currentJob.Name]) { $jobAssignments[$currentJob.Name] = @() }
                    $jobAssignments[$currentJob.Name] += "Team: $($itemObject.ToString()) (Email: $($itemObject.Mail))"
                } catch {
                    "Failed to add Team {0} (Email: {1}) to job {2}: {3}" -f $itemObject.ToString(), $itemObject.Mail, $currentJob.Name, $_.Exception.Message | timelog
                }
            }
            "GroupMailbox" {
                if ($itemObject -and $itemObject -is [Veeam.Archiver.PowerShell.Model.VBOOrganizationGroup]) {
                    $item = New-VBOBackupItem -Group $itemObject -GroupMailbox:$withGroupMailbox
                    try {
                        $groupIdentifier = if ($itemObject.DisplayName) { $itemObject.DisplayName } else { $itemObject.GroupName }
                        Add-VBOBackupItem -Job $currentJob -BackupItem $item -ErrorAction Stop
                        "Added GroupMailbox {0} (GroupName: {1}) to job {2} (Mailbox included: {3})" -f $groupIdentifier, $itemObject.GroupName, $currentJob.Name, $withGroupMailbox | timelog
                        $objCount++
                        if (-not $jobAssignments[$currentJob.Name]) { $jobAssignments[$currentJob.Name] = @() }
                        $jobAssignments[$currentJob.Name] += "GroupMailbox: $groupIdentifier (GroupName: $($itemObject.GroupName))"
                    } catch {
                        "Failed to add GroupMailbox {0} (GroupName: {1}) to job {2}: {3}" -f $itemObject.ToString(), $itemObject.GroupName, $currentJob.Name, $_.Exception.Message | timelog
                    }
                } else {
                    "Skipping invalid GroupMailbox object: $($itemObject)" | timelog
                }
            }
            "Site" {
                $siteItem = New-VBOBackupItem -Site $itemObject
                try {
                    Add-VBOBackupItem -Job $currentJob -BackupItem $siteItem -ErrorAction Stop
                    "Added Site {0} (URL: {1}) to job {2}" -f $itemObject.ToString(), $itemObject.URL, $currentJob.Name | timelog
                    $objCount++
                    if (-not $jobAssignments[$currentJob.Name]) { $jobAssignments[$currentJob.Name] = @() }
                    $jobAssignments[$currentJob.Name] += "Site: $($itemObject.ToString()) (URL: $($itemObject.URL))"
                } catch {
                    "Failed to add Site {0} (URL: {1}) to job {2}: {3}" -f $itemObject.ToString(), $itemObject.URL, $currentJob.Name, $_.Exception.Message | timelog
                }
                
                if (!$limitServiceTo) {
                    # Add matching Team if exists
                    $team = $teams | Where-Object { ($_.Mail -split "@")[0] -eq ([uri]$itemObject.URL).Segments[-1] }
                    if ($team) {
                        $teamItem = New-VBOBackupItem -Team $team -TeamsChats:$withTeamsChats
                        try {
                            Add-VBOBackupItem -Job $currentJob -BackupItem $teamItem -ErrorAction Stop
                            "Added matching Team {0} (Email: {1}) to job {2}" -f $team.ToString(), $team.Mail, $currentJob.Name | timelog
                            $objCount++
                            if (-not $jobAssignments[$currentJob.Name]) { $jobAssignments[$currentJob.Name] = @() }
                            $jobAssignments[$currentJob.Name] += "Team: $($team.ToString()) (Email: $($team.Mail))"
                            # Remove from $teams to avoid re-adding later
                            $teams = $teams | Where-Object { $_.Mail -ne $team.Mail }
                        } catch {
                            "Failed to add Team {0} (Email: {1}) to job {2}: {3}" -f $team.ToString(), $team.Mail, $currentJob.Name, $_.Exception.Message | timelog
                        }
                    }
                    # Add matching GroupMailbox if exists
                    $siteUrl = ([uri]$itemObject.URL).AbsolutePath.TrimEnd('/')
                    $siteName = $itemObject.ToString().Replace(" ", "").ToLower()
                    $group = $groupMailboxes | Where-Object { 
                        $groupSiteUrl = if ($_.SiteUrl) { ([uri]$_.SiteUrl).AbsolutePath.TrimEnd('/') } else { $null }
                        $groupSiteUrl -eq $siteUrl
                    } | Select-Object -First 1
                    if (-not $group) {
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
                            $groupIdentifier = if ($group.DisplayName) { $group.DisplayName } else { $group.GroupName }
                            Add-VBOBackupItem -Job $currentJob -BackupItem $groupItem -ErrorAction Stop
                            "Added matching GroupMailbox {0} (GroupName: {1}) to job {2} (Mailbox included: {3})" -f $groupIdentifier, $group.GroupName, $currentJob.Name, $withGroupMailbox | timelog
                            $objCount++
                            if (-not $jobAssignments[$currentJob.Name]) { $jobAssignments[$currentJob.Name] = @() }
                            $jobAssignments[$currentJob.Name] += "GroupMailbox: $groupIdentifier (GroupName: $($group.GroupName))"
                            # Remove from $groupMailboxes to avoid re-adding later
                            $groupMailboxes = $groupMailboxes | Where-Object { $_.GroupName -ne $group.GroupName }
                        } catch {
                            "Failed to add GroupMailbox {0} (GroupName: {1}) to job {2}: {3}" -f $groupIdentifier, $group.GroupName, $currentJob.Name, $_.Exception.Message | timelog
                        }
                    } else {
                        "No matching GroupMailbox found for site {0} (URL: {1}, Name: {2}) in job {3}" -f $itemObject.ToString(), $itemObject.URL, $siteName, $currentJob.Name | timelog
                        if ($groupMailboxes.Count -gt 0) {
                            "Available GroupMailboxes: {0}" -f ($groupMailboxes | ForEach-Object { 
                                $normalizedName = if ($_.DisplayName) { $_.DisplayName.Replace(" ", "").Replace("M365Group", "").ToLower() } else { "N/A" }
                                "$($_.DisplayName) (Normalized Name: $normalizedName, SiteUrl: $($_.SiteUrl), GroupName: $($_.GroupName))"
                            } | Join-String -Separator ", ") | timelog
                        }
                    }
                }
            }
        }
    }
}

END {
    "Summary of object assignments to jobs:" | timelog
    foreach ($jobName in $jobAssignments.Keys) {
        "Job $jobName contains:" | timelog
        $jobAssignments[$jobName] | ForEach-Object { "  - $_" | timelog }
    }
    Stop-Transcript
}