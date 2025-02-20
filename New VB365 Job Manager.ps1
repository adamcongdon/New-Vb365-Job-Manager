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
        foreach($j in $jobs){
            # print which job we are working on
            Write-Host "Working on job: $($j.Name)"
            $items = Get-VBOBackupItem -Job $j 
            foreach($i in $items){
                if($i.Mailbox -eq $true){
                    if($i.GroupMailbox -eq $false){
                        #print the group mailbox we are working on
                        Write-Host "Working on Group Mailbox: $($i.Group)"
                        Set-VBOBackupItem -BackupItem $i -GroupMailbox:$true
                        Add-VBOBackupItem -Job $j -BackupItem $i
                    }
                }
            }
        }