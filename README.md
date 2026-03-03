GroupToContactSyncThing
A PowerShell script used to synchronize groups from a source Microsoft Entra tenant into mail contacts in a target Entra tenant.
This is commonly used for cross‑tenant GAL visibility, where groups in one tenant need to appear as contacts in another.

Usage
PowerShell.\GroupToContactSyncThing.ps1 `  -SourceTenantDomain <source tenant domain> `  -TargetTenantDomain <target tenant domain> `  -LogRoot <log folder path> `  -StampCustomAttribute <extension attribute number> `  -GroupRecipientTypeDetails <array of group types> `  [-Schedule]Show more lines

Parameters
-SourceTenantDomain
The primary domain name of the source Entra tenant where the groups reside.
Example:
Plain Textcontoso.onmicrosoft.comShow more lines

-TargetTenantDomain
The primary domain name of the target Entra tenant where the groups will be created as mail contacts.
Example:
Plain Textfabrikam.onmicrosoft.comShow more lines

-LogRoot
The local folder path where logs and transcripts will be written.
Example:
Plain TextC:\Logs\GroupToContactSyncShow more lines

The script will create timestamped subfolders for each run.


-StampCustomAttribute
The number of the CustomAttribute (1–15) used on the target mail contact to store the
ExternalDirectoryObjectId of the source group.
This stamp is used to:

Ensure idempotency
Reliably match source groups to existing target contacts
Prevent duplicate creation

Example:
PowerShell-StampCustomAttribute 15Show more lines

-GroupRecipientTypeDetails
An array of group recipient types to be synchronized from the source tenant.
Common values include:

MailUniversalDistributionGroup
MailUniversalSecurityGroup
GroupMailbox (Microsoft 365 groups)

Example:
PowerShell-GroupRecipientTypeDetails @(  "MailUniversalDistributionGroup",  "MailUniversalSecurityGroup",  "GroupMailbox")Show more lines

-Schedule
(Optional switch)
When specified, the script will create a scheduled task on the local machine to run the sync automatically.

This switch should typically be used once, during initial setup.

Example:
PowerShell-Sc
