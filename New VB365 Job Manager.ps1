# New VB365 Job Manager

<#
Synopsis

This script will check for objects that are not protected by a Veeam Backup for Microsoft 365 job and sort them into jobs based on the number of objects in the job.

If the job has reached the limit of objects, a new job will be created and the objects will be added to the new job.

First we will check for and sort Teams and Sharepoint
#>




<##
Variables:



1. Get each org and look for unprocessed Group Mailbox
2. Get jobs with Templated Name and check count of items in the job
3. If there is room in the job, then add new items to that job.
3.1 if no room in the job, create a new job and add the items to the new job.

#>
# Variables:
$itemsPerJob = 2
$jobPrefix = "JM-GroupMailboxes-"
$repository = "storj_cloud"


# #FUNCTIONS
# function CreateNewJob{
#     param(
#         [Parameter(Mandatory=$true)]
#         [string]$name,
#         [Parameter(Mandatory=$true)]
#         [string]$org,
#         [Parameter(Mandatory=$true)]
#         [string]$repo,
#         [Parameter(Mandatory=$true)]
#         [string]$items
#     )
#     $o = Get-VBOOrganization -Name $org
#     $rep = Get-VBORepository -Name $repo
#     $job = Add-VBOJob -Organization $o -Name $name -Repository $rep -SelectedItems $items
#     return $job
# }
# get all orgs
$organizations = Get-VBOOrganization
$repo = Get-VBORepository -Name $repository

# Loop through each org
foreach ($org in $organizations){
    #append org suffix to job
    $jobName = $jobPrefix + $org.Name + "-"
    # Get all the Group Mailboxes that are not protected
    $NotProtectedGroupMailboxes = Get-VBOOrganizationGroup -Organization $org -NotInJob

    #if there are no Group Mailboxes that are not protected, move on to the next org
    if ($NotProtectedGroupMailboxes.Count -eq 0){
        continue
    }

    # Get the jobs and number of Groups per job
    $jobs = Get-VBOJob # -Organization $org

    #check job names against $jobName
    $jobs = $jobs | Where-Object {$_.Name -like "$jobName*"}
    
    # if no jobs are found, create a new job
    if ($jobs.Count -eq 0){
        $createNewJobFirst = $true
    }
    # else use the existing job with the least amount of groups
    else{
        foreach($j in $jobs){
            $itemCount = Get-VBOBackupItem -Job $j | Measure-Object
            if($itemCount.Count -lt $itemsPerJob){
                $job = $j
            }
        }
    }

    # Add the Group Mailboxes to the job
    

    foreach($group in $NotProtectedGroupMailboxes){
        $newItem = New-VBOBackupItem -Group $group -GroupMailbox

        if($createNewJobFirst){
            $newJobName = $jobName + "1"
            $jobCount = 1
            #$job = CreateNewJob -name $newJobName -org $org -repo $repository -items $newItem
            $job = Add-VBOJob -Organization $org -Name $newJobName -Repository $repo -SelectedItems $newItem
            $createNewJobFirst = $false
        }
        $itemsInJob = Get-VBOBackupItem -Job $job

        if($itemsInJob.Count -lt $itemsPerJob){
            #Set-VBOJob -Job $job -SelectedItems $newItem
            Add-VBOBackupItem -Job $job -BackupItem $newItem
        }
        else{
            #get repository from current job:
            #$repository = Get-VBORepository -Job $job
            # create new job and set itemsInJob to 0
            $jobCount++
            $job = Add-VBOJob -Organization $org -Name "$jobName$jobCount" -Repository $repo -SelectedItems $newItem
        }
    }
    while($itemsInJob.Count -lt $itemsPerJob){


    }




}