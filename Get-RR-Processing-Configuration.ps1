
<#
This script prompts the user to search for a room or equipment mailbox.  Once the search is
narrowed down to a single mailbox, the user can proceed to showing how the mailbox is 
configured to process bookings from the users.

The user is also shown the permissions assigned to the mailbox's calendar folder.

No settings are changed using this script. It's only used to view the configuration.

REQUIRED: A PowerShell session with Exchange Online.

#>


$RR = $null

Do {
    
    $RR = $null
    $RRCount = ""
    
    ## Prompt for user input
    Write-Host
    $UserInput = Read-Host -Prompt "Search for a room/resource (RR) name (Wildcard * Supported)"

    $RR = (Get-mailbox $UserInput -filter { RecipientTypeDetails -eq "RoomMailbox" -or RecipientTypeDetails -eq "EquipmentMailbox" } | select Name)

    $RRCount = $RR.Name.Count

    If ($RRCount -gt "1") {
        
        Write-Host
        Write-Host -ForegroundColor Green "Too many RRs found. Narrow your search."
        Write-Host -ForegroundColor Green "RRs Found: $($RRCount)"
        Write-Host -ForegroundColor Green "---------------------------" 

        ForEach ($R in $RR) {
            
            Write-Host -ForegroundColor Green "$($R.Name)"
            
        }

        Write-Host
    }

} Until ($RRCount -eq "1")


Write-Host
Write-Host -ForegroundColor Green "RR Found: $($RR.Name)"
Write-Host
$GetRR = ""
$GetRR = Read-Host -Prompt "Do you want the calendar processing configuration for this RR? [y|n]" 



If ( $GetRR -eq "Y") {

    $RRConf = Get-CalendarProcessing $RR.Name

    $AllBookIn = $RRConf.AllBookInPolicy
    $AllRequestIn = $RRConf.AllRequestInPolicy
    $AllRequestOut = $RRConf.AllRequestOutOfPolicy

    $ListBookIn = $RRConf.BookInPolicy
    $ListRequestIn = $RRConf.RequestInPolicy
    $ListRequestOut = $RRConf.RequestOutOfPolicy
    $ListDelegate = $RRConf.ResourceDelegates

    $BookWinDays = $RRConf.BookingWindowInDays
    $MaxDuration = $RRConf.MaximumDurationInMinutes

    If ($BookWinDays -ne "370") {
        
        Write-Host
        Write-Host -ForegroundColor Red -BackgroundColor Black "Warning: This RR is miss-configured."
        Write-Host -ForegroundColor Red -BackgroundColor Black "The Max Booking Lead Time in Days is set to $($BookWinDays)."
        Write-Host -ForegroundColor Red -BackgroundColor Black "The correct value should be 370, and can be set in the Exchange Admin Console in Office 365."
    }

    If ($MaxDuration -ne "1440") {
        
        Write-Host
        Write-Host -ForegroundColor Red -BackgroundColor Black "Warning: This RR is miss-configured."
        Write-Host -ForegroundColor Red -BackgroundColor Black "The Max Duration in Minutes is set to $($MaxDuration)."
        Write-Host -ForegroundColor Red -BackgroundColor Black "The correct value should be 1440, and can be set in the Exchange Admin Console in Office 365."
    }


    If ( ($AllBookIn -eq $true) -and ($AllRequestIn -eq $false) -and ($AllRequestOut -eq $false) ) {
        
        Write-Host
        Write-Host -ForegroundColor Green "Anyone can book $($RR.Name) via auto-processing."
        Write-Host -ForegroundColor Green "No delegate approval required."
        Write-Host


    }

    If ( ($AllBookIn -eq $false) -and ($AllRequestIn -eq $true) -and ($AllRequestOut -eq $false) ) {
    
        Write-Host
        Write-Host -ForegroundColor Green "Only the following list of people can book $($RR.Name) via auto-processing."
        Write-Host -ForegroundColor Green "Everyone else requires delegate approval."
        Write-Host
        Write-Host -ForegroundColor Green "Users"
        Write-Host -ForegroundColor Green "------------"
        
        ForEach ( $User in $ListBookIn ){
            
            $DisplayName = Get-Recipient $User | Select DisplayName
            Write-Host -ForegroundColor Green "$($DisplayName.DisplayName)"


        }

        Write-Host 
        Write-Host -ForegroundColor Green "Delegates"
        Write-Host -ForegroundColor Green "------------"
        
        ForEach ( $User in $ListDelegate ){
            
            $DisplayName = Get-Recipient $User | Select DisplayName
            Write-Host -ForegroundColor Green "$($DisplayName.DisplayName)"


        }

    
    
    }

    If ( ($AllBookIn -eq $false) -and ($AllRequestIn -eq $false) -and ($AllRequestOut -eq $false) -and ($ListDelegate.count -gt 0)  ) {
    
        Write-Host
        Write-Host -ForegroundColor Green "Only the following delegates can book $($RR.Name) via auto-processing."
        Wirte-Host -ForegroundColor Green "Everyone else is denied."
        Write-Host
        Write-Host -ForegroundColor Green "Delegates"
        Write-Host -ForegroundColor Green "------------"

        ForEach ( $User in $ListDelegate ){
            
            $DisplayName = Get-Recipient $User | Select DisplayName
            Write-Host -ForegroundColor Green "$($DisplayName.DisplayName)"


        }
    
    
    }

    If ( ($AllBookIn -eq $false) -and ($AllRequestIn -eq $false) -and ($AllRequestOut -eq $false) -and ($ListDelegate.count -eq 0) ) {
    
        Write-Host
        Write-Host -ForegroundColor Green "Only the following list of people can book $($RR.Name) via auto-processing."
        Write-Host -ForegroundColor Green "All others are denied." 
        Write-Host -ForegroundColor Green "There are no delegates."
        Write-Host
        Write-Host -ForegroundColor Green "Users"
        Write-Host -ForegroundColor Green "------------"
    
        ForEach ( $User in $ListBookIn ){
            
            $DisplayName = Get-Recipient $User | Select DisplayName
            Write-Host -ForegroundColor Green "$($DisplayName.DisplayName)"


        }

    
    }


    $FolderName = $RR.Name + ":\Calendar"
    $FolderPerms = Get-MailboxFolderPermission $FolderName

    $OriginalColor = $Host.ui.rawui.foregroundcolor

    Write-Host
    Write-Host
    Write-Host -ForegroundColor Green "Calendar Folder Permissions: $($RR.Name)"
    
    $Host.ui.rawui.foregroundcolor = "Green"
    
    $FolderPerms | FT 

    $Host.ui.rawui.foregroundcolor = $OriginalColor

    Write-Host

} Else { Write-Host; Write-Host -ForegroundColor Green "GOOD BYE!"; Write-Host; Exit}


