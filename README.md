# SharePoint_File_Tree_Export
Requires PowerShell 7+

Open PS7 as admin and install the PNP Module
Install-Module PnP.PowerShell -Scope CurrentUser -Force

Register the certificate and app on the tenancy
Register-PnPEntraIDApp -ApplicationName "SPO-StorageReport" `
    -Tenant <client>.onmicrosoft.com `
    -Store CurrentUser `
    -SharePointApplicationPermissions "Sites.Read.All"

Make a note of the thumbprint & client ID

Now you can run the script with your chosen paramters
.\SP_File_Tree.ps1 `
    -AdminUrl https://<client>-admin.sharepoint.com `
    -ClientId <client-id> `
    -Tenant <client>.onmicrosoft.com `
    -CertificateThumbprint <thumbprint> `
    -SiteUrl https://<client>.sharepoint.com/sites/<a-known-site> `
    -AlsoExportCsv
