#requires -version 4
<#
.SYNOPSIS
  This will gather all in place hold policies are report it's members to a CSV

.DESCRIPTION
  The script will sign into O365 then query the In-place hold policies.  For each policy found, it will the gather the members found in the SourceMailboxes parameter, query Active
  Directory for additional user information, which then is added into a custom object and placed in an array.  At the end of the script the array will then be export to a CSV file.

.PARAMETER <Parameter_Name>
  None

.INPUTS
  None

.OUTPUTS
  C:\Temp\HoldPolicy(currentDate).csv

.NOTES
  Version:        1.0
  Author:         John Kovolski
  Creation Date:  Oct 10, 2016
  Purpose/Change: Initial script development

.EXAMPLE
  None
    
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

# Generates file for credentials
Get-Credential | Export-Clixml "C:\scripts\credentials.xml"
$credpath = "C:\scripts\credentials.xml"
$Cred = Import-Clixml $credpath

#Import Modules & Snap-ins
Import-Module MSOnline

# Connect to Exchange Online and the Security and Compliance Center
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell-liveid -Credential $cred -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid -Credential $Cred -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber -DisableNameChecking

Connect-msolservice -Credential $cred

#----------------------------------------------------------[Declarations]----------------------------------------------------------

$file = "C:\temp\holdpolicy-$(Get-Date -f yyyy-MM-dd-HHmmss).csv"

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Write-Host 'Gathering Hold Policy Information ...'

$checkPolicy = @(Get-MailboxSearch)
$policyinfo = @()

    
foreach ($policy in $checkPolicy)
 {
    
# Gather info on each In Place hold policy         
   $check = Get-MailboxSearch -Identity $policy.Name 

   Write-Host -ForegroundColor Cyan "Gathered $($check.Name) policy members.  Continuing..."
           
   $MbxList = $check.SourceMailboxes

   ForEach ($Mbx in $MbxList) 
    {

# Create custom object for output used later
      $rc = New-Object PSObject
      $rc | Add-Member -type NoteProperty -name PolicyName -Value $policy.Name
      $rc | Add-Member -type NoteProperty -name InPlaceHoldEnabled -Value $policy.InPlaceHoldEnabled
      $rc | Add-Member -type NoteProperty -name ItemHoldPeriod -Value $policy.ItemHoldPeriod

# For each member of the In Place hold policy, get AD info
      $checkmb = Get-Mailbox -Identity $mbx -ErrorAction SilentlyContinue | select UserPrincipalName
      $upn = $checkmb.UserPrincipalName
      $user = Get-ADUser -filter {UserPrincipalName -eq $upn} -Properties * | Select DisplayName, Mail, SamAccountName

      $rc | Add-Member -type NoteProperty -name MemberName -Value $user.DisplayName
      $rc | Add-Member -type NoteProperty -name Email -Value $user.Mail
      $rc | Add-Member -type NoteProperty -name SamAccountName -Value $user.SamAccountName

      $policyinfo += $rc
    } 

   
 }           

$policyInfo | select PolicyName, MemberName, EMail,SamAccountName, InPlaceHoldEnabled, ItemHoldPeriod  | Export-Csv $file -NoTypeInformation
Write-Host -ForegroundColor Green "Process completed.  Please check review $($file)."