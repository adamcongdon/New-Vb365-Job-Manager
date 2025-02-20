# 

<#
Synopsis
This script will check for jobs processing mailboxes and add the GroupMailbox if not being processed.

#>

<##

1. Get all jobs
2. Check jobs for mailboxes
3. If Mailbox is backed up, add in Group Mailbox

#>


$jobs = Get-VBOJob # -Organization $org
foreach ($j in $jobs) {
    # print which job we are working on
    #Write-Host "Checking job to ensure : $($j.Name)"
    $items = Get-VBOBackupItem -Job $j 
    $mailboxes = $items | Where-Object { $_.Mailbox -eq $true }
    if($mailboxes.Count -eq 0){
        #Write-Host "No mailboxes in job: $($j.Name)"
        continue
    }
    else{
        Write-Host "Mailboxes found in job: $($j.Name)"
        Write-Host "Checking to ensure Group Mailboxes are included"
    }
    foreach ($i in $mailboxes) {
        if ($i.Mailbox -eq $true) {
            if ($i.GroupMailbox -eq $false) {
                #print the group mailbox we are working on
                Write-Host "Working on Group Mailbox: $($i.Group)"
                Set-VBOBackupItem -BackupItem $i -GroupMailbox:$true
                Add-VBOBackupItem -Job $j -BackupItem $i
            }
        }
    }
}



<## Need to work on adding missing groups to new/existing jobs
Something like:
	1	Get groups that are not in a job
	2	Add these groups to jobs:
	a	Job type 1: Mailbox, Archive, Group Mail 
	b	Job type 2: OneDrive
	c	Job type 3: Sharepoint, Teams (should be collected by existing job manager script)

    #>




<##



1. Get each org and look for unprocessed Group Mailbox
2. Get jobs with Templated Name and check count of items in the job
3. If there is room in the job, then add new items to that job.
3.1 if no room in the job, create a new job and add the items to the new job.

#>
# Variables:
$itemsPerJob = 5
$jobPrefixMailBox = "JM-Mailboxes-"
$jobPrefixOneDrive = "JM-OneDrive-"
$repository = "storj_cloud"


# get all orgs
$organizations = Get-VBOOrganization
$repo = Get-VBORepository -Name $repository

# Loop through each org
foreach ($org in $organizations) {
    $notProtectedGroups = Get-VBOOrganizationGroup -Organization $org -NotInJob
    
    if ($notProtectedGroups.Count -eq 0) {
        Write-Host "No groups to process for $($org.Name)"
        continue
    }

    # Get the jobs and number of Groups per job
        $mbJobs = $jobs | Where-Object { $_.Name -like "$jobNameMB*" }
        #$odJobs = $jobs | Where-Object { $_.Name -like "$odJobName*" }
    
        $jobCount = $mbJobs.count + 1
    foreach($group in $notProtectedGroups){
        Write-Host "Working on group: $($group.DisplayName)"
        $newMbItem = New-VBOBackupItem -Group $group -GroupMailbox -Mailbox -ArchiveMailbox
        $newOdItem = New-VBOBackupItem -Group $group -OneDrive
        $needNewJob = $false
        $jobs = Get-VBOJob  -Organization $org

        $mailboxAdded = $false
        $oneDriveAdded = $false
        foreach($job in $jobs){
            #write-host "Working on job: $($job.Name)"
            $items = Get-VBOBackupItem -Job $job
            if($mailboxAdded -eq $true -and $oneDriveAdded -eq $true){
                Write-Host "Both items added, moving to next group"
                $needNewJob = $false
                break
            }
            if($items.count -ge $itemsPerJob){
                Write-Host "Job is full, moving to next job. Job: $($job.Name)"
                $needNewJob = $true
                continue
            }
            else{
                $needNewJob = $false
                $mailboxes = $items | Where-Object { $_.Mailbox -eq $true }
                $oneDrive = $items | Where-Object { $_.OneDrive -eq $true }
                if ($mailboxes.count -gt 0) {
                    Write-Host "Adding Mailbox Item to $($job.Name)"
                    Add-VBOBackupItem -Job $job -BackupItem $newMbItem
                    $mailboxAdded = $true
                }
                if ($oneDrive.count -gt 0) {
                    Write-Host "Adding OneDrive Item to $($job.Name)"
                    Add-VBOBackupItem -Job $job -BackupItem $newOdItem
                    $oneDriveAdded = $true
                }
            }
        }
        if($needNewJob){
            # add logging here 
            Write-Host "Creating new job for $($group.Name)"

                $newJobName = $jobNameMB + $jobCount
                $newOdJobName = $odJobName + $jobCount
                #$jobCount = 1
                #$job = CreateNewJob -name $newJobName -org $org -repo $repository -items $newItem
                $job = Add-VBOJob -Organization $org -Name $newJobName -Repository $repo -SelectedItems $newMbItem
                $odJob = Add-VBOJob -Organization $org -Name $newOdJobName -Repository $repo -SelectedItems $newOdItem
                $jobCount++
                Write-Host "MB Job Created: $($job.Name)"
                Write-Host "OD Job Created: $($odJob.Name)"
        }
    }
    
    # foreach ($job in $jobs) {
    #     Write-Host "Working on job: $($job.Name)"
    #     $items = Get-VBOBackupItem -Job $job
    #     if ($items.Count -gt $itemsPerJob) {
    #         continue
    #     }
    #     $mailboxes = $items | Where-Object { $_.Mailbox -eq $true }
    #     $oneDrive = $items | Where-Object { $_.OneDrive -eq $true }

    #     if ($mailboxes.count -gt 0) {
    #         foreach ($group in $notProtectedGroups) {
    #             $newItem = New-VBOBackupItem -Group $group -GroupMailbox -Mailbox -ArchiveMailbox
    #             Add-VBOBackupItem -Job $job -BackupItem $newItem
    #         }
    #     }
    #     if ($oneDrive.count -gt 0) {
    #         foreach ($group in $notProtectedGroups) {
    #             $newItem = New-VBOBackupItem -Group $group -OneDrive
    #             Add-VBOBackupItem -Job $job -BackupItem $newItem
    #         }
    #     }

    # }
    # # all existing jobs should be filled up to max items. New jobs will be created if needed.
    # $remainingGroups = Get-VBOOrganizationGroup -Organization $org -NotInJob 
    # if ($remainingGroups.count -eq 0) {
    #     continue
    # }
    # else {
    #     # add groups to new jobs
    #     $jobNameMB = $jobPrefixMailBox + $org.Name + "-"
    #     $odJobName = $jobPrefixOneDrive + $org.Name + "-"
    #     #check job names against $jobNameMB
    #     $mbJobs = $jobs | Where-Object { $_.Name -like "$jobNameMB*" }
    #     $odJobs = $jobs | Where-Object { $_.Name -like "$odJobName*" }
    
    #     # if no jobs are found, create a new job
    #     if ($mbJobs.Count -eq 0) {
    #         $createNewJobFirst = $true
    #     }
        
    #     foreach ($group in $remainingGroups) {
    #         $newItem = New-VBOBackupItem -Group $group -GroupMailbox -Mailbox -ArchiveMailbox
    #         $OneDriveItem = New-VBOBackupItem -Group $group -OneDrive
        
    #         if ($createNewJobFirst) {
    #             $newJobName = $jobNameMB + "1"
    #             $newOdJobName = $odJobName + "1"
    #             $jobCount = 1
    #             #$job = CreateNewJob -name $newJobName -org $org -repo $repository -items $newItem
    #             $job = Add-VBOJob -Organization $org -Name $newJobName -Repository $repo -SelectedItems $newItem
    #             $odJob = Add-VBOJob -Organization $org -Name $newOdJobName -Repository $repo -SelectedItems $OneDriveItem
    #             $createNewJobFirst = $false
    #         }
           
    #         $itemsInJob = Get-VBOBackupItem -Job $job
    #         #$itemsInOdJob = Get-VBOBackupItem -Job $odJob
        
    #         if ($itemsInJob.Count -lt $itemsPerJob) {
    #             #Set-VBOJob -Job $job -SelectedItems $newItem
    #             Add-VBOBackupItem -Job $job -BackupItem $newItem
    #             Add-VBOBackupItem -Job $odJob -BackupItem $OneDriveItem
    #         }
    #         else {
    #             #get repository from current job:
    #             #$repository = Get-VBORepository -Job $job
    #             # create new job and set itemsInJob to 0
    #             $jobCount++
    #             $job = Add-VBOJob -Organization $org -Name "$jobNameMB$jobCount" -Repository $repo -SelectedItems $newItem
    #             $odJob = Add-VBOJob -Organization $org -Name "$odJobName$jobCount" -Repository $repo -SelectedItems $OneDriveItem
    #         }
    #     }

    # }

    # # Add the Group Mailboxes to the job
    

    



}