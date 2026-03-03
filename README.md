# GroupToContactSyncThing

A PowerShell script used to synchronize **groups from a source Microsoft Entra tenant** into **mail contacts in a target Entra tenant**.

This is commonly used for **cross-tenant GAL visibility**, where groups in one tenant need to appear as contacts in another.

**NOTE** : at the top of the script there are 4 variables defined as user defined, these are not passed in the script PARAMS section, you must define the attributes that will be populated (other than mail) that will be used when creating vs. updating an existing contact\mailcontact.

---

## Usage

```powershell
.\GroupToContactSyncThing.ps1   -SourceTenantDomain <source tenant domain>   -TargetTenantDomain <target tenant domain>   -LogRoot <log folder path>   -StampCustomAttribute <extension attribute number>   -GroupRecipientTypeDetails <array of group types>   [-Schedule]
```

---

## Parameters

### `-SourceTenantDomain`
The **primary domain name** of the source Entra tenant where the groups reside.

Example:
```text
contoso.onmicrosoft.com
```

---

### `-TargetTenantDomain`
The **primary domain name** of the target Entra tenant where the groups will be created as mail contacts.

Example:
```text
fabrikam.onmicrosoft.com
```

---

### `-LogRoot`
The **local folder path** where logs and transcripts will be written.

Example:
```text
C:\Logs\GroupToContactSync
```

The script creates timestamped subfolders for each run.

---

### `-StampCustomAttribute`
The **CustomAttribute number (1–15)** used on the target mail contact to store the
`ExternalDirectoryObjectId` of the source group.

This stamp is used to:
- Ensure uniqueness
- Reliably match source groups to existing target contacts
- Prevent duplicate creation

Example:
```powershell
-StampCustomAttribute 15
```

---

### `-GroupRecipientTypeDetails`
An **array of group recipient types** to synchronize from the source tenant.

Common values include:
- `MailUniversalDistributionGroup`
- `MailUniversalSecurityGroup`

Example:
```powershell
-GroupRecipientTypeDetails @(
  "MailUniversalDistributionGroup",
  "MailUniversalSecurityGroup"
)
```

---

### `-Schedule`
Optional switch.

When specified, the script creates a **scheduled task** on the local machine to run the sync automatically.

This switch should typically be used **once**, during initial setup.

Example:
```powershell
-Schedule
```
