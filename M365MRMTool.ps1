function Show-Menu {
    Write-Host "============================="
    Write-Host "  M365 MRM Policy Tool"
    Write-Host "============================="
    Write-Host "1. List Existing MRM Policies"
    Write-Host "2. Apply MRM Policy to Mailbox"
    Write-Host "3. View Current MRM Policy on Mailbox"
    Write-Host "4. Check Mailbox Usage"
    Write-Host "5. Exit"
    Write-Host "============================="
}

function List-MRMPolicies {
    Write-Host "Listing existing MRM policies with details..."
    
    # Get all MRM policies
    $policies = Get-RetentionPolicy
    
    foreach ($policy in $policies) {
        Write-Host "`nPolicy Name: $($policy.Name)"
        
        # Get the tags associated with this policy
        $tags = $policy.RetentionPolicyTagLinks | ForEach-Object {
            Get-RetentionPolicyTag -Identity $_
        }
        
        # Display details for each tag
        foreach ($tag in $tags) {
            Write-Host "  Tag Name: $($tag.Name)"
            
            # Inspect the AgeLimitForRetention property
            $ageLimit = $tag.AgeLimitForRetention
            if ($ageLimit) {
                $ageLimitRaw = $ageLimit.ToString()
                Write-Host "  Retention Period: $ageLimitRaw"
            } else {
                Write-Host "  Retention Period: Not Set"
            }
            
            Write-Host "  Action: $($tag.RetentionAction)"
            Write-Host "  Type: $($tag.Type)"
            Write-Host ""
        }
    }

    # Pause to allow the user to view the list before returning to the menu
    Read-Host "Press Enter to return to the menu"
}

function Apply-MRMPolicyToMailbox {
    param(
        [string]$Mailbox,
        [string]$PolicyName
    )
    Write-Host "Setting MRM policy to null for mailbox '$Mailbox'..."
    Set-Mailbox -Identity $Mailbox -RetentionPolicy $null
    
    Write-Host "Applying MRM policy '$PolicyName' to mailbox '$Mailbox'..."
    Set-Mailbox -Identity $Mailbox -RetentionPolicy $PolicyName

    Write-Host "Starting Managed Folder Assistant for mailbox '$Mailbox'..."
    Start-ManagedFolderAssistant -Identity $Mailbox
    
    Write-Host "MRM policy '$PolicyName' has been applied to mailbox '$Mailbox'."
    Read-Host "Press Enter to return to the menu"
}

function View-CurrentMRMPolicy {
    param(
        [string]$Mailbox
    )
    # Check if the archive mailbox is enabled
    $mailboxDetails = Get-Mailbox -Identity $Mailbox
    if ($mailboxDetails.ArchiveStatus -eq "Active") {
        Write-Host "Archive mailbox is enabled for '$Mailbox'."
        
        # Get current retention policy applied to the mailbox
        $retentionPolicy = Get-Mailbox -Identity $Mailbox | Select-Object -ExpandProperty RetentionPolicy
        if ($retentionPolicy) {
            Write-Host "Current MRM policy applied to mailbox '$Mailbox': $retentionPolicy"
        } else {
            Write-Host "No MRM policy is currently applied to mailbox '$Mailbox'."
        }
    } else {
        Write-Host "Archive mailbox is not enabled for '$Mailbox'."
        $enableArchive = Read-Host "Do you want to enable the archive mailbox? (Y/N)"
        if ($enableArchive -eq "Y" -or $enableArchive -eq "y") {
            Write-Host "Enabling archive mailbox for '$Mailbox'..."
            Enable-Mailbox -Identity $Mailbox -Archive
            Write-Host "Archive mailbox has been enabled for '$Mailbox'."
        } else {
            Write-Host "Archive mailbox will not be enabled."
        }
    }

    # Pause to allow the user to view the details before returning to the menu
    Read-Host "Press Enter to return to the menu"
}

function Check-MailboxUsage {
    param(
        [string]$Mailbox
    )
    do {
        Write-Host "Checking mailbox usage for '$Mailbox'..."

        # Get mailbox usage details
        $mailboxUsage = Get-MailboxStatistics -Identity $Mailbox
        $archiveUsage = if ($mailboxUsage.ArchiveStatus -eq "Active") {
            Get-MailboxStatistics -Identity $Mailbox -Archive | Select-Object -ExpandProperty TotalItemSize
        } else {
            "Archive mailbox is not enabled."
        }

        Write-Host "Mailbox Size: $($mailboxUsage.TotalItemSize)"
        Write-Host "Mailbox Item Count: $($mailboxUsage.ItemCount)"
        Write-Host "Last Logon Time: $($mailboxUsage.LastLogonTime)"
        Write-Host "Last Logoff Time: $($mailboxUsage.LastLogoffTime)"
        Write-Host "Archive Mailbox Size: $archiveUsage"

        # Prompt to recheck the same mailbox or return to the main menu
        $recheck = Read-Host "Do you want to recheck the same mailbox? (Y/N)"
        if ($recheck -eq "N" -or $recheck -eq "n") {
            return
        }
    } while ($recheck -eq "Y" -or $recheck -eq "y")
}

function Main {
    try {
        # Sign in as the global admin
        Connect-ExchangeOnline -UserPrincipalName (Read-Host "Enter Global Admin UPN") -ShowProgress $true
        
        do {
            cls
            Show-Menu
            $choice = Read-Host "Select an option (1-5)"
            switch ($choice) {
                1 {
                    List-MRMPolicies
                }
                2 {
                    $mailbox = Read-Host "Enter mailbox identity"
                    $policyName = Read-Host "Enter MRM policy name to apply"
                    Apply-MRMPolicyToMailbox -Mailbox $mailbox -PolicyName $policyName
                }
                3 {
                    $mailbox = Read-Host "Enter mailbox identity"
                    View-CurrentMRMPolicy -Mailbox $mailbox
                }
                4 {
                    $mailbox = Read-Host "Enter mailbox identity"
                    Check-MailboxUsage -Mailbox $mailbox
                }
                5 {
                    Write-Host "Exiting..."
                    break
                }
                default {
                    Write-Host "Invalid choice, please select an option from 1 to 5."
                }
            }
        } while ($choice -ne 5)
    }
    catch {
        Write-Host "An error occurred during sign-in. Please check your credentials and try again."
    }
    finally {
        # Disconnect after operations
        Disconnect-ExchangeOnline -Confirm:$false
    }
}

# Start the tool
Main
