<#
THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR 
IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

	NAME:	GroupToContactSyncThing.ps1
	AUTHOR:	Darryl Kegg
	DATE:	9 December, 2023
	EMAIL:	dkegg@microsoft.com
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [Parameter(Mandatory=$true)]
  [string]$SourceTenantDomain = "source.onmicrosoft.com",

  [Parameter(Mandatory=$true)]
  [string]$TargetTenantDomain = "target.onmicrosoft.com",

  [string]$LogRoot = "C:\Logs\GroupToContactSyncThing",

  [ValidateSet(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15)]
  [int]$StampCustomAttribute = 15,

  [string[]]$GroupRecipientTypeDetails = @("MailUniversalDistributionGroup","MailUniversalSecurityGroup"),

  [switch]$WhatIfOnly,

  [Parameter(Position = 0, Mandatory = $false, HelpMessage = "Creates 'scheduled task' in Windows Scheduler that can be adjusted in the future to perform synchronization cycle(s) at particular time interval")]
  [Switch]$Schedule
)

# Switch to directory where script is located
Push-Location (split-path -parent $MyInvocation.MyCommand.Definition)

if ($schedule)
{
              $taskname = "GroupToContactSyncThing"
              schtasks.exe /create /RL HIGHEST /NP /sc MINUTE /MO 5 /st 06:00 /tn "$taskname" /tr "$PSHOME\powershell.exe -c '. ''$($myinvocation.mycommand.definition)'''"
              exit
}

# -------------------------
# USER-DEFINED ATTRIBUTES
# -------------------------
$CreateContactAttrs = @{
  Company    = "Thing1"
  Department = "Cross-Tenant Syncthing"
  Title      = "Group copy"
}

$CreateMailContactAttrs = @{
  HiddenFromAddressListsEnabled = $false
  CustomAttribute1              = "GroupDuper"
}

$UpdateContactAttrs = @{
  Company    = "Thing1"
  Department = "Cross-Tenant Syncthing"
}

$UpdateMailContactAttrs = @{
  HiddenFromAddressListsEnabled = $false
  CustomAttribute1              = "GroupDuper"
}

# -------------------------
# LOGGING SETUP
# -------------------------
$ScriptName = "GroupToContactSyncThing"
$RunStamp   = Get-Date -Format "yyyyMMdd-HHmmss"
$RunFolder  = Join-Path $LogRoot $RunStamp

New-Item -ItemType Directory -Path $RunFolder -Force | Out-Null

$startTime = Get-Date

$TranscriptPath = Join-Path $RunFolder "$ScriptName-$RunStamp.transcript.txt"
Start-Transcript -Path $TranscriptPath -Force | Out-Null  

$ChangeLogPath = Join-Path $RunFolder "$ScriptName-$RunStamp.changes.csv"
"Timestamp,Action,ObjectType,SourceStamp,GroupDisplayName,GroupPrimarySmtp,TargetIdentity,ChangedFields,Message" |
  Out-File -FilePath $ChangeLogPath -Encoding UTF8 -Force

function New-UniqueRecipientStrings {
  param(
    [Parameter(Mandatory=$true)][string]$BaseName,     # group display name
    [Parameter(Mandatory=$true)][string]$BaseAlias,    # derived alias
    [Parameter(Mandatory=$true)][string]$StampValue    # group object id/guid string
  )

  # Use a short deterministic suffix so it’s stable across runs
  $suffix = $StampValue.Replace("-", "")
  if ($suffix.Length -gt 10) { $suffix = $suffix.Substring($suffix.Length - 10) }

  # Candidate Name: FriendlyName + stable suffix
  $nameCandidate  = "$BaseName [$suffix]"
  # Candidate Alias: grp_alias + stable suffix
  $aliasCandidate = ("$BaseAlias" + "_" + $suffix).ToLower()

  # sanitize Alias to valid characters (simple safe set)
  $aliasCandidate = $aliasCandidate -replace "[^a-z0-9_]", ""
  if ($aliasCandidate.Length -gt 64) { $aliasCandidate = $aliasCandidate.Substring(0,64) }

  # Ensure Name is unique (recipient identity namespace)
  $n = $nameCandidate
  $i = 1
  while (Get-Recipient -ResultSize 1 -Filter ("Name -eq '{0}'" -f $n) -ErrorAction SilentlyContinue) {
    $n = "$nameCandidate-$i"
    $i++
  }

  # Ensure Alias is unique (alias/mailnickname collisions)
  $a = $aliasCandidate
  $j = 1
  while (Get-Recipient -ResultSize 1 -Filter ("Alias -eq '{0}'" -f $a) -ErrorAction SilentlyContinue) {
    $a = "$aliasCandidate$j"
    if ($a.Length -gt 64) { $a = $a.Substring(0,64) }
    $j++
  }

  return [pscustomobject]@{
    UniqueName  = $n
    UniqueAlias = $a
    Suffix      = $suffix
  }
}

function Write-ChangeLog {
  param(
    [Parameter(Mandatory)] [ValidateSet("CREATE","UPDATE","SKIP","ERROR")] [string]$Action,
    [Parameter(Mandatory)] [string]$ObjectType,
    [string]$SourceStamp,
    [string]$GroupDisplayName,
    [string]$GroupPrimarySmtp,
    [string]$TargetIdentity = "",
    [string]$ChangedFields = "",
    [string]$Message = ""
  )

  if ($null -eq $SourceStamp)      { $SourceStamp = "" }
  if ($null -eq $GroupDisplayName) { $GroupDisplayName = "" }
  if ($null -eq $GroupPrimarySmtp) { $GroupPrimarySmtp = "" }
  if ($null -eq $TargetIdentity)   { $TargetIdentity = "" }
  if ($null -eq $ChangedFields)    { $ChangedFields = "" }
  if ($null -eq $Message)          { $Message = "" }

  $line = ('"{0}","{1}","{2}","{3}","{4}","{5}","{6}","{7}","{8}"' -f
    (Get-Date -Format "o"),
    $Action,
    $ObjectType,
    $SourceStamp.Replace('"','""'),
    $GroupDisplayName.Replace('"','""'),
    $GroupPrimarySmtp.Replace('"','""'),
    $TargetIdentity.Replace('"','""'),
    $ChangedFields.Replace('"','""'),
    $Message.Replace('"','""')
  )

  Add-Content -Path $ChangeLogPath -Value $line
}

# functions go here
function Connect-ExoTenant {
  param([string]$Org)
  Write-Host "Connecting to Exchange Online: $Org" -ForegroundColor Cyan
  Connect-ExchangeOnline -Organization $Org -ShowBanner:$false
}

function Disconnect-ExoTenant {
  Disconnect-ExchangeOnline -Confirm:$false | Out-Null
}

function Get-MailEnabledGroups {
  param([string[]]$TypeDetails)
  $all = Get-Recipient -ResultSize Unlimited
  $all | Where-Object { $TypeDetails -contains $_.RecipientTypeDetails }
}

function Get-StampValue {
  param($Group)
  if ($Group.ExternalDirectoryObjectId) { return $Group.ExternalDirectoryObjectId.ToString() }
  if ($Group.Guid) { return $Group.Guid.ToString() }
  return $null
}

function Get-TargetContactByStampOrMail {
  param(
    [string]$StampAttrName,
    [string]$StampValue,
    [string]$ExternalSmtp
  )

  if ($StampValue) {
    $filter = "$StampAttrName -eq '$StampValue'"
    $c = Get-MailContact -ResultSize Unlimited -Filter $filter -ErrorAction SilentlyContinue
    if ($c) { return $c | Select-Object -First 1 }
  }

  $smtpFilter = "ExternalEmailAddress -eq 'SMTP:$ExternalSmtp'"
  $c2 = Get-MailContact -ResultSize Unlimited -Filter $smtpFilter -ErrorAction SilentlyContinue
  if ($c2) { return $c2 | Select-Object -First 1 }

  return $null
}

function Compare-AndBuildSplat {
  param(
    [hashtable]$Desired,
    $CurrentObject
  )

  $toSet = @{}
  foreach ($k in $Desired.Keys) {
    $desiredVal = $Desired[$k]
    $currentVal = $CurrentObject.$k

    $cv = if ($null -eq $currentVal) { "" } elseif ($currentVal -is [Array]) { ($currentVal -join ";") } else { [string]$currentVal }
    $dv = if ($null -eq $desiredVal) { "" } elseif ($desiredVal -is [Array]) { ($desiredVal -join ";") } else { [string]$desiredVal }

    if ($cv -ne $dv) { $toSet[$k] = $desiredVal }
  }

  return $toSet
}

# -------------------------
# MAIN SCRIPT BLOCK HERE (I don't like MAIN{})
# -------------------------
$ErrorActionPreference = "Stop"
$stampAttrName = "CustomAttribute$StampCustomAttribute"

try {
  # Connect to the source tenant and read all my defined group types
  Connect-ExoTenant -Org $SourceTenantDomain
  $groups = Get-MailEnabledGroups -TypeDetails $GroupRecipientTypeDetails
  Write-Host "Source groups found: $($groups.Count)" -ForegroundColor Green
}
catch {
  Write-ChangeLog -Action "ERROR" -ObjectType "SourceTenant" -Message $_.Exception.Message
  throw
}
finally {
    # not using the source anymore, so gracefully disconnect
    Disconnect-ExoTenant
}

try {
  # Connect to the target tenant and use this connection for the rest of the activities
  Connect-ExoTenant -Org $TargetTenantDomain

  foreach ($g in $groups) {
    $groupSmtp = [string]$g.PrimarySmtpAddress
    $groupName = [string]$g.DisplayName

    if (-not $groupSmtp) { continue }

    $stampValue = Get-StampValue -Group $g
    if (-not $stampValue) {
      Write-ChangeLog -Action "ERROR" -ObjectType "MailContact" -SourceStamp "" `
        -GroupDisplayName $groupName -GroupPrimarySmtp $groupSmtp `
        -Message "No stamp value available (ExternalDirectoryObjectId/Guid missing)."
      continue
    }

    $contact = $null

    try {
      $contact = Get-TargetContactByStampOrMail -StampAttrName $stampAttrName -StampValue $stampValue -ExternalSmtp $groupSmtp

      $aliasBase = if ($g.Alias) { $g.Alias } else { $groupSmtp.Split("@")[0] }
      $alias = ("grp_" + $aliasBase) -replace "[^a-zA-Z0-9_]", ""
      if ($alias.Length -gt 64) { $alias = $alias.Substring(0,64) }

      if (-not $contact) {
        # CREATE
        if ($PSCmdlet.ShouldProcess("TargetTenant:$TargetTenantDomain", "Create MailContact for group '$groupName' ($groupSmtp)")) {
          if (-not $WhatIfOnly) {
            $baseAlias = $alias  
            $u = New-UniqueRecipientStrings -BaseName $groupName -BaseAlias $baseAlias -StampValue $stampValue

            $new = New-MailContact -Name $u.UniqueName `
            -DisplayName $groupName `
            -Alias $u.UniqueAlias `
            -ExternalEmailAddress $groupSmtp

            $createMail = @{}
            $createMail += $CreateMailContactAttrs
            $createMail[$stampAttrName] = $stampValue 
            if ($createMail.Count -gt 0) { Set-MailContact -Identity $new.Identity @createMail }

            if ($CreateContactAttrs.Count -gt 0) { Set-Contact -Identity $new.Identity @CreateContactAttrs }

            Write-Host "Created: $groupName -> $groupSmtp ($stampAttrName=$stampValue)" -ForegroundColor Green

            $changedFields = @("New-MailContact", $stampAttrName, "ExternalEmailAddress")
            $changedFields += $CreateMailContactAttrs.Keys
            $changedFields += $CreateContactAttrs.Keys

            Write-ChangeLog -Action "CREATE" -ObjectType "MailContact" -SourceStamp $stampValue `
              -GroupDisplayName $groupName -GroupPrimarySmtp $groupSmtp -TargetIdentity $new.Identity `
              -ChangedFields ($changedFields -join ";")
          }
          else {
            Write-ChangeLog -Action "CREATE" -ObjectType "MailContact" -SourceStamp $stampValue `
              -GroupDisplayName $groupName -GroupPrimarySmtp $groupSmtp `
              -ChangedFields "WhatIfOnly" -Message "Would create contact"
          }
        }
      }
      else {
        # UPDATE instead
        $desiredMail = @{}
        $desiredMail += $UpdateMailContactAttrs
        $desiredMail[$stampAttrName] = $stampValue

        $currentExternal = [string]$contact.ExternalEmailAddress 
        $desiredExternal = "SMTP:$groupSmtp"
        if ($currentExternal -ne $desiredExternal) {
          $desiredMail["ExternalEmailAddress"] = $groupSmtp
        }

        $mailDelta = Compare-AndBuildSplat -Desired $desiredMail -CurrentObject $contact

        $contactObj = Get-Contact -Identity $contact.Identity
        $contactDelta = Compare-AndBuildSplat -Desired $UpdateContactAttrs -CurrentObject $contactObj

        if (($mailDelta.Count -gt 0) -or ($contactDelta.Count -gt 0)) {
          if ($PSCmdlet.ShouldProcess("TargetTenant:$TargetTenantDomain", "Update MailContact '$($contact.DisplayName)' for group '$groupName'")) {
            if (-not $WhatIfOnly) {
              if ($mailDelta.Count -gt 0)    { Set-MailContact -Identity $contact.Identity @mailDelta }  
              if ($contactDelta.Count -gt 0) { Set-Contact     -Identity $contact.Identity @contactDelta } 

              $changed = @()
              if ($mailDelta.Count -gt 0)    { $changed += ($mailDelta.Keys    | ForEach-Object { "Mail:$($_)" }) }
              if ($contactDelta.Count -gt 0) { $changed += ($contactDelta.Keys | ForEach-Object { "Contact:$($_)" }) }

              Write-Host "Updated: $($contact.DisplayName) (deltas: $($changed -join '; '))" -ForegroundColor Yellow

              Write-ChangeLog -Action "UPDATE" -ObjectType "MailContact" -SourceStamp $stampValue `
                -GroupDisplayName $groupName -GroupPrimarySmtp $groupSmtp -TargetIdentity $contact.Identity `
                -ChangedFields ($changed -join ";")
            }
            else {
              $would = @()
              $would += $mailDelta.Keys
              $would += $contactDelta.Keys
              Write-ChangeLog -Action "UPDATE" -ObjectType "MailContact" -SourceStamp $stampValue `
                -GroupDisplayName $groupName -GroupPrimarySmtp $groupSmtp -TargetIdentity $contact.Identity `
                -ChangedFields "WhatIfOnly" -Message ("Would update: " + ($would -join ";"))
            }
          }
        }
        else {
          Write-Host "No change: $($contact.DisplayName)" -ForegroundColor DarkGray
          Write-ChangeLog -Action "SKIP" -ObjectType "MailContact" -SourceStamp $stampValue `
            -GroupDisplayName $groupName -GroupPrimarySmtp $groupSmtp -TargetIdentity $contact.Identity `
            -ChangedFields "No changes"
        }
      }
    }
    catch {
      $targetId = ""
      if ($null -ne $contact -and $null -ne $contact.Identity) {
        $targetId = [string]$contact.Identity
      }

      Write-Host "ERROR processing group '$groupName': $($_.Exception.Message)" -ForegroundColor Red
      Write-ChangeLog -Action "ERROR" -ObjectType "MailContact" -SourceStamp $stampValue `
        -GroupDisplayName $groupName -GroupPrimarySmtp $groupSmtp `
        -TargetIdentity $targetId -ChangedFields "" -Message $_.Exception.Message
      continue
    }
  }
}
catch {
  Write-ChangeLog -Action "ERROR" -ObjectType "TargetTenant" -Message $_.Exception.Message
  throw
}
finally {
  Disconnect-ExoTenant

    $endTime = Get-Date
    
    $duration = New-TimeSpan -Start $startTime -End $endTime
    
    Write-Host "Total execution time: $($duration.Hours)h $($duration.Minutes)m $($duration. Seconds)s"

  Stop-Transcript | Out-Null 

  Write-Host "Run folder: $RunFolder" -ForegroundColor Cyan
  Write-Host "Change log : $ChangeLogPath" -ForegroundColor Cyan
  Write-Host "Transcript : $TranscriptPath" -ForegroundColor Cyan

}
