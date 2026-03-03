# GroupToContactSyncThing
Powershell script that can be used to sync groups as contacts from one Entra tenant to another

Usage :

  -SourceTenantDomain <your source tenant> 
  
  -TargetTenantDomain <your target tenant> 
  
  -logroot <path to a folder for logging eg. C:\Logs\GroupToContactSync> 
  
  -StampCustomAttribute <number of the extension attribute to use on the target contact for stamping the externalDirectoryObjectID from the source >
  
  -GroupRecipientTypeDetails <array of group types to convert from the source>
  
  -Schedule : use this switch one time to create a scheduled task on the machine if you want it to be automated
  

  
