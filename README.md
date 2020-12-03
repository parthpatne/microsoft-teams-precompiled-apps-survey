# Action Package deployment PowerShell script

This package contains the PowerShell scripts for deploying action packages to ActionPlatform

## Step 1 : Creating ActionPackage.zip
Download the compiled action package from GitHub <Enter Url>.<br/>
Update the icon or other resources in the package.<br/>
Update the name and urls in actionManifest.json.<br/>
Create zip file of the compiled content.<br/>


## Step 2 : Upload the package zip to ActionPlatform using PowerShell script
Open a PowerShell console.<br/>
Change directory to the directory containing UploadActionPackage.ps1 file.<br/> 
Run following command.<br/>
```UploadActionPackage.ps1 -PackageZipFilePath <ActionPackageZipFilePath> [-TeamsAppDownloadDirectoryPath <TeamsAppDownloadDirectoryPath>] [-LogLevel <LogLevel>] [-LogDirectoryPath <LogDirectoryPath>] [-Endpoint <Endpoint>] [-AccessToken <AccessToken>]```

### PackageZipFilePath (The only mandatory parameter)
User needs to provide path to the compiled action package zip.

### TeamsAppDownloadDirectoryPath (optional parameter)
Directory where the script downloads final Teams App manifest zip. If this parameter is not provided, app zip will be downloaded in working directory.

### LogLevel (optional parameter)
This parameter can be used to set console logging level. By default the log level is set to `info`.
- **error** - Just error logs will be shown
- **warning** - Error and warning messages will be shown
- **info** - Beside status messages, informative logs will be shown
- **debug** - More debug logs will be shown
- **none** - No logs will be shown

### LogDirectoryPath (optional parameter)
Directory where the script stores the log file. If this parameter is not provided, logs will be stored in working directory.

### Endpoint (optional parameter)
Set to "https://actions.office365.com" by default. 

### AccessToken (optional parameter)
If this script fails to acquire the token due to MSAL PowerShell module installation or any other issues, then user can manually acquire the token and provide it as input to this script.
