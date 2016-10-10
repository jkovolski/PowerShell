#requires -version 4
<#
.SYNOPSIS
  This will gather all in place hold policies are report it's members to a CSV

.DESCRIPTION
  The script will sign into O365 then query the In-place hold policies.  For each policy found, it will the gather the members found in the SourceMailboxes parameter and add it to an array.  At the end
  of the script the array will then be export to a CSV file.

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

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
  #None 
)

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

#Any Global Declarations go here

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Write-Host 'Gathering Hold Policy Information ...'


Try {
      $checkPolicy = @(Get-MailboxSearch)
      $policyinfo = @()
      
        foreach ($policy in $checkPolicy)
        {
          $userlist = $null
          $check = Get-MailboxSearch -Identity $policy.Name 
          $policyinfo += $check

          Write-Host -ForegroundColor Cyan "Gathered $($check.Name) policy members.  Continuing..."
         
         }           
    }
    
    Catch {
      Write-Warning -ForegroundColor Red "Error:$($check.Name) with error $($_.Exception)"
      Break
      }


 $policyInfo | select Name, @{Name='Members'; Expression={$_.sourcemailboxes -join ","}}  | Export-Csv "C:\temp\holdpolicy-$(Get-Date -f yyyy-MM-dd-HHmmss).csv" -NoTypeInformation