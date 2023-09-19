# sharepoint-server-manager
PowerShell scripts to manage SharePoint Server farm

## Files
| File | Role |
| - | - |
| **_Helpers.ps1** | Useful methods |
| **Config.LAB.json** | Configuration file with tenant and SharePoint URL, app registration and CSV separator |
| **Create-AppService.ps1** | Create New-SPSubscriptionSettingsServiceApplication and New-SPAppManagementServiceApplication |
| **Create-ManagedMetadataService.ps1** | Create New-SPMetadataServiceApplication |
| **Create-SearchService.ps1** | Create New-SPMetadataServiceApplication |
| **Reset-SearchTopology.ps1** | Recreate the search topology and move the SearchData directory |
| **Create-UserProfileService.ps1** | Create New-SPProfileServiceApplication |
| **Test-SendMail.ps1** | Test the SMTP configuration by sending a new mail |


## Prerequisities

## Configuration

Create a new configuration file based on `Config.LAB.json` or edit this one.

When executing the scripts, the code-name of the configuration should be passed as an argument:

~~~powershell
.\Test-SendMail.ps1 -Env LAB
.\Test-SendMail.ps1 -Env PROD
~~~

Sample of LAB configuration is:

~~~json
{
	"WebApplications": ["https://your-webapp.local", "https://your-mysites.local"],
	"ADDomain": "dev.local",
	"ADPath": "CN=Users,DC=dev,DC=local",
	"UserProfileServiceAppName": "User Profile Service Application",
	"UserProfileServiceProxyName": "User Profile Service Application Proxy",
	"AppManagementServiceAppName": "App Management Service Application",
	"AppManagementServiceProxyName": "App Management Service Application Proxy",
	"ManagedMetadataServiceAppName": "Managed Metadata Service Application",
	"ManagedMetadataServiceProxyName": "Managed Metadata Service Application Proxy",
	"AppSubscriptionServiceAppName": "Subscriptions Settings Service Application",
	"SearchServiceAppName": "Search Service Application",
	"SearchServiceProxyName": "Search Service Application Proxy",
	"AppPoolAccount": "DEV\\dev_shp_Srv",
	"AppPoolName": "SharePoint Service Applications",
	"UserProfileDB": "SP_Profile",
	"UserProfileSyncDB": "SP_Sync",
	"UserProfileSocialDB": "SP_Social",
	"AppManagementDB": "SP_AppServiceApp",
	"ManagedMetadataDB": "SP_MMS",
	"SearchDB": "SP_Search",
	"Servers": ["sp2019-dev"],
	"SearchServer": "aphp-win-1",
	"ServerDataDrive": "D",
	"SearchDataDirectory": "SearchData",	
	"DatabaseServer": "sp2019-dev\\SHAREPOINT"
}
~~~

A directory called `Logs` is automatically created when executing scripts.