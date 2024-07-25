# Remove App Packages

This script remove App Packages from Teams

You need to create an app registration in your tenant. 

- Client Flow need to be enabled
- Add the permission AppCatalog.ReadWrite.All

You need to be global Admin in your tenant to be able to delete AppPackages

## Run

You can run from commandline:
>> node index.js clientId TenantId AppName (Version)

The script will filter on the AppName and optional on Version.

For each package you can decide to delete or not. Or give All if you are brave.