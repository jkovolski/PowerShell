#requires -version 4
<#
.SYNOPSIS
  Import a CSV containing shared mailboxes, the autoreply message, and forwarding address if applicable.  
  Then update the shared mailbox settings accordingly

.DESCRIPTION
  This script will update the shared mailbox settings with a new autoreply message for both internal and external communication.  
  If there is a required forwarding email address it will be added to the shared mailbox.

.PARAMETER <Parameter_Name>
  None.

.INPUTS
  None.

.OUTPUTS Log File
  The script log file stored in C:\Temp\SharedMailboxUpdates-$(Get-Date -f yyyy-MM-dd-HHmmss).log

.NOTES
  Version:        1.0
  Author:         John Kovolski
  Creation Date:  9/27/2016
  Purpose/Change: Initial script development

.EXAMPLE
  <Example explanation goes here>
  
  <Example goes here. Repeat this attribute for more than one example>
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------
<#[CmdletBinding(SupportsShouldProcess=$true)]
Param (
    [parameter(Mandatory=$true, HelpMessage="Enter the path of the CSV file with Shared Mailbox settings.")]
    [ValidateNotNullOrEmpty()]
 #   [ValidateScript({Test-path $_ -isValid })]
    [string]$File

)#>
#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

#Update this varible with the file path of the shared mailbox settings that you wish to update.
$file = "c:\scripts\sharedmailbox.csv"

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
$sScriptVersion = '1.0'

#Log File Info
$sLogPath = 'C:\Temp'
$sLogName = "SharedMailboxUpdates-$(Get-Date -f yyyy-MM-dd-HHmmss).log"
$sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName

#-----------------------------------------------------------[Functions]------------------------------------------------------------

Function LogWrite {

   Param ([string]$logstring)
   Add-content $sLogFile -value "$(Get-Date) : $logstring"

}


## FUNCTION - LogErrorHalt
## Writes the error to the log, sends an email, and stops the script.
## Example: If ($Error) { LogErrorHalt "$Error.exception" }

Function LogErrorHalt {

    Param ([string]$ErrorString)
    Add-Content $sLogFile -value "$(Get-Date) : ===============   Start Script Error   ==============="
    Add-content $sLogFile -value "$(Get-Date) : $ErrorString"
    Add-Content $sLogFile -value "$(Get-Date) : ===============    End Script Error    ==============="
    Add-content $sLogFile -value "$(Get-Date) : An error has occurred. Stopping the script."
    Add-Content $sLogFile -value "$(Get-Date) : =============== End Shared Mailbox Update Log ==============="
     
 #   Send-MailMessage -SmtpServer $SMTPrelay -From $EmailFrom -To $EmailTo -Subject $ScriptHaltEmailSubject -Body $ScriptHaltEmailBody -Attachments $EmailAttachment -Priority High

    Exit

}


## FUNCTION - LogErrorContinue
## Writes the error to the log and continues with the script.
## Example: If ($Error) { LogErrorContinue "$Error.exception" }

Function LogErrorContinue {

    Param ([string]$ErrorString)
    Add-Content $sLogFile -value "$(Get-Date) : ===============   Start Script Error   ==============="
    Add-content $sLogFile -value "$(Get-Date) : $ErrorString"
    Add-Content $sLogFile -value "$(Get-Date) : ===============    End Script Error    ==============="

}


## FUNCTION - SendErrorLog
## When errors have been logged, this function is called to send an email with the log file attached.
## Example: If ($ErrorCount -gt 0) { SendErrorLog "$sLogFile" }


#-----------------------------------------------------------[Execution]------------------------------------------------------------
## Test the log file path and create the log file. 
## Halt the script on error.

If(!(Test-Path -Path $sLogPath )){ New-Item -ItemType directory -Path $sLogPath -Force }
If ($Error) { LogErrorHalt "$Error.exception" }

New-Item $sLogFile -ItemType File -Force
If ($Error) { LogErrorHalt "$Error.exception" }

## Write first line of the log using the LogWrite function.

LogWrite "=============== Start Shared Mailbox Log ==============="

$sharedMailboxes = Import-Csv $file

LogWrite "Imported content successfully."


foreach ($sharedMailbox in $sharedMailboxes)
{
    Set-MailboxAutoReplyConfiguration -Identity $sharedMailbox.mailbox -AutoReplyState enabled -ExternalMessage $sharedMailbox.message -InternalMessage $sharedMailbox.message 
    Logwrite "Adding Autoreply to $($sharedMailbox.mailbox)."
    Write-Host -ForegroundColor Green "Updated $($sharedMailbox.mailbox)."

    If ($Error) { LogErrorHalt "$Error.exception" }


    if (($sharedMailbox.forward))
    {
        Set-Mailbox  $sharedMailbox.mailbox -ForwardingSmtpAddress $sharedMailbox.forward -DeliverToMailboxAndForward $true 
        Logwrite "Adding forwarding for $($sharedMailbox.mailbox).  Emails will be forwaded to $($sharedmailbox.forward)."
       
    }
        else
        {
            Logwrite "There is no forwarding address for $($sharedMailbox.mailbox).  Moving to next mailbox."
            continue
        }
   Logwrite "All shared mailboxes have been updated. "
   Write-Host -ForegroundColor Green "Completed Shared Mailbox updates. Please review logfile for more details. "
}

