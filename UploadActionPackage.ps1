param(
    [Parameter(Mandatory = $true, HelpMessage = "Action package zip file local path")]
    [ValidateScript( { Test-Path $_ -PathType leaf })]
    [string]$PackageZipFilePath,

    [Parameter(Mandatory = $false, HelpMessage = "Teams app zip download directory path")]    
    [string]$TeamsAppDownloadDirectoryPath,

    [Parameter(Mandatory = $false, HelpMessage = "Log Level")]
    [ValidateSet("Status", "Info", "Debug", "None")] 
    [string]$LogLevel = "Info",

    [Parameter(Mandatory = $false, HelpMessage = "Log directory path")]
    [string]$LogDirectoryPath,

    [Parameter(Mandatory = $false, HelpMessage = "Action platform endpoint")]    
    [string]$Endpoint = "https://actions.office365.com",    

    [Parameter(Mandatory = $false, HelpMessage = "AccessToken acquired manually, if automated token acquisition fails")]
    [string]$AccessToken
)


# Create Correlation Id for this session
$RequestCorrelationId = [guid]::NewGuid().ToString()

# Maximum time API calls must take
$MAX_API_TIMEOUT = 30 # seconds

# Maximum number of retries in API calls
$MAX_API_RETRIES = 3

# Minimum API retry interval in exponential retry pattern
$MIN_RETRY_DELAY = 2  #seconds

# Http error codes that need to be retried
$RETRYABLE_ERROR_CODES = @(429 <# TooManyRequests #>, 408 <# RequestTimeout #>, 502 <# BadGateway #>, 503 <# ServiceUnavailable #>, 504 <# GatewayTimeout #>)

# Maximum number of retries to monitor status url
$MAX_MONITORING_RETRIES = 30


#region logging
 
$LogLevelMap = @{            
    'None'    = 0
    'Success' = 1
    'Error'   = 2
    'Warning' = 3
    'Info'    = 4
    'Debug'   = 5
}

function Initialize-Logger {
    param (
        [Parameter(Mandatory = $false)]
        [string]$LogDirectoryPath        
    )

    if ([string]::IsNullOrWhiteSpace($LogDirectoryPath) -or !(Test-Path -Path $LogDirectoryPath -PathType Container -ErrorAction Stop)) {
        $LogDirectoryPath = "$(Get-Location)\ActionPackageLogs"
    }

    $TimeStamp = Get-Date -Format o | ForEach-Object { $_ -replace ":", "." } 
    $LogFilePath = "$LogDirectoryPath\$TimeStamp.log"

    $LogFile = New-Item -Path $LogFilePath -Force -ItemType File 
    Write-Host "Log file path: $LogFilePath"

    return $LogFilePath
}

function Write-Log {
    param (
        [Parameter(Mandatory = $false)]
        [Alias("LogLevel")]
        [ValidateSet("Success", "Error", "Warning", "Info", "Debug")] 
        [string]$StatementLogLevel = "Info",

        [Parameter(Mandatory = $true)]
        [string]$Message        
    )

    # Generate date string to log
    $DateTimeToLog = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 

    "$DateTimeToLog $StatementLogLevel $Message" | Out-File -FilePath $LogFilePath -Append    

    if ($LogLevelMap[$LogLevel] -lt $LogLevelMap[$StatementLogLevel]) {
        return
    }

    switch ($StatementLogLevel) {
        'Success' {  
            Write-Host $Message -ForegroundColor Green
        }
        'Error' {  
            Write-Host $Message -ForegroundColor Red
        }
        'Warning' {  
            Write-Host $Message -ForegroundColor Yellow
        }
        'Info' {  
            Write-Host $Message
        }
        'Debug' {  
            Write-Host $Message
        }        
    }
}

function Write-ErrorLog {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message   
    )
    
    Write-Log -LogLevel "Error" -Message $Message
}

function Write-WarningLog {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message   
    )
    
    Write-Log -LogLevel "Warning" -Message $Message
}

function Write-InfoLog {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message   
    )
    
    Write-Log -LogLevel "Info" -Message $Message
}

function Write-SuccessLog {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message   
    )
    
    Write-Log -LogLevel "Success" -Message $Message 
}

function Write-DebugLog {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message   
    )
    
    Write-Log -LogLevel "Debug" -Message $Message
}

#endregion logging


#region utils

function ExitOnError {
    param (
        [Parameter(Mandatory = $false)]
        [bool]$IsError = $true,

        [Parameter(Mandatory = $true)]
        [String]$OnErrorMessage
    )
    
    if ($IsError) {
        Write-ErrorLog $OnErrorMessage
        throw $OnErrorMessage
    }
}

function ExecuteAPI {
    param (
        [Parameter(Mandatory = $true)]
        [String]$Method,
        
        [Parameter(Mandatory = $true)]
        [String]$Uri,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$Headers,

        [Parameter(Mandatory = $false)]
        [System.Object]$Body, 

        [Parameter(Mandatory = $false)]
        [int]$RetryAttempt = 0,
        
        [Parameter(Mandatory = $false)]
        [String]$OnErrorMessage   
    )

    try {
        Write-DebugLog "Executing API: Method: $Method, Uri: $Uri, Body: $Body"
        $response = Invoke-RestMethod -Uri $Uri -Method $Method -Headers $Headers -Body $Body -TimeoutSec $MAX_API_TIMEOUT -ErrorAction Stop
        return $response
    }
    catch {
        Write-DebugLog "API execution failed. $PSItem"

        if ([bool]($_.Exception.PSobject.Properties.name -match "Response")) {

            Write-ErrorLog "API response error. StatusCode: $($_.Exception.Response.StatusCode.value__), StatusDescription: $($_.Exception.Response.StatusDescription)"

            $StatusCode = $_.Exception.Response.StatusCode.value__
            if ($RETRYABLE_ERROR_CODES.Contains($StatusCode)) {
                Write-DebugLog "$StatusCode is a retryable status code. Current RetryAttempt# $RetryAttempt"

                if ($RetryAttempt -le $MAX_API_RETRIES) {
                    $Delay = [math]::Pow($MIN_RETRY_DELAY, $RetryAttempt)
                    Write-DebugLog "Retrying after $Delay"
                    Start-Sleep -Seconds $Delay
                    return ExecuteAPI -Method $Method -Uri $Uri -Headers $Headers -Body $Body -RetryAttempt ($RetryAttempt + 1) -OnErrorMessage $OnErrorMessage
                }
                else {
                    Write-ErrorLog "API failed after $RetryAttempt retries. "
                }
            }
        }

        if (!([string]::IsNullOrWhiteSpace($OnErrorMessage))) {
            Write-ErrorLog $OnErrorMessage
        }

        throw
    }
}

function ExecuteActionAPI {
    param (        
        [Parameter(Mandatory = $true)]
        [String]$Method,

        [Parameter(Mandatory = $true)]
        [String]$Uri,

        [Parameter(Mandatory = $false)]
        [System.Object]$Body = $null,

        [Parameter(Mandatory = $false)]
        [String]$OnErrorMessage        
    )    
    $Headers = GetActionHeaders $AccessToken
    return ExecuteAPI -Method $Method -Uri $Uri -Headers $Headers -Body $Body -OnErrorMessage $OnErrorMessage
}

function GetActionHeaders {
    return @{            
        'Authorization'        = "Bearer $AccessToken"
        'Content-Type'         = "application/json"
        'RequestCorrelationId' = $RequestCorrelationId
        'Accept-Encoding'      = "gzip, deflate"
    } 
}

#endregion utils


#region ActionService APIs

function ValidateOrAcquireToken {
    if ([string]::IsNullOrWhiteSpace($AccessToken)) {
        if (!(Get-Module -ListAvailable -Name MSAL.PS)) {            
            Write-DebugLog "MSAL.PS module is not already installed"
            Write-InfoLog "Need to install MSAL.PS module for authentication purpose. Updating Nuget Package and PowerShellGet Module for the same"
    
            # ## Update Nuget Package and PowerShellGet Module
            Install-PackageProvider NuGet -Force -Scope CurrentUser 
            Install-Module PowerShellGet -Force -Scope CurrentUser -AllowClobber
    
            # ## In a new PowerShell process, install the MSAL.PS Module. Restart PowerShell console if this fails.
            Write-InfoLog "Installing the MSAL.PS Module. Restart PowerShell console if this fails."
        
            &(Get-Process -Id $pid).Path -Command { Install-Module MSAL.PS -Scope CurrentUser }
            Import-Module MSAL.PS
        
            if (!(Get-Module -ListAvailable -Name MSAL.PS)) {
                ExitOnError -OnErrorMessage "Not able to find MSAL.PS module. Please restart PowerShell console and try again, or provide the AccessToken as parameter to this script"
            }
        }
        else {
            Write-DebugLog "MSAL.PS module is already installed."
        }
    
        $Scope = "$($Endpoint)/ActionPackage.ReadWrite.All" 

        $ClientId = "cac88df7-3599-49cf-9465-867b9eee33cf"
        $RedirectUri = "urn:ietf:wg:oauth:2.0:oob"
        $Authority = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
    
        Write-DebugLog "Fetching API token for endpoint: $Endpoint, scope: $Scope..."
        Write-InfoLog "Please login to your AAD account when prompted ..."

        $authResponse = Get-MsalToken -Scope $Scope -ClientId $ClientId -RedirectUri $RedirectUri -Authority $Authority -Prompt 'SelectAccount'
    
        if ($null -eq $authResponse -or [string]::IsNullOrWhiteSpace($authResponse.AccessToken)) {
            ExitOnError -OnErrorMessage "Token acquisition failed. Please try again later"
        }
    
        Write-DebugLog "Token acquisition successful"
        return $authResponse.AccessToken
    }
    else {
        Write-InfoLog "Using AccessToken provided as input parameter"    
        return $AccessToken
    }
}

function GetActionPackageUploadUrl {
    $Uri = "$Endpoint/v1/actionPackages/zipUploadUrl"
    
    Write-InfoLog "Fetching ActionPackage zip upload url ..."
    $ErrorMessage = "Failed to get zip upload url!"
    $response = ExecuteActionAPI -Method "Get" -Uri $Uri -OnErrorMessage $ErrorMessage
    
    $PackageZipUploadUrl = $response.url

    ExitOnError -IsError ([string]::IsNullOrWhiteSpace($PackageZipUploadUrl)) -OnErrorMessage $ErrorMessage

    Write-DebugLog "PackageZipUploadUrl: $PackageZipUploadUrl"
    return $PackageZipUploadUrl
}

function UploadPackageZipToBlob {
    param (        
        [Parameter(Mandatory = $true)]
        [String]$PackageZipUploadUrl,

        [Parameter(Mandatory = $true)]
        [String]$PackageZipFilePath
    )
    
    $Headers = @{    
        'Content-Type'   = "application/zip"
        'x-ms-blob-type' = "BlockBlob"
    }

    Write-InfoLog "Uploading action package zip ..."
    $FileBytes = [System.IO.File]::ReadAllBytes($PackageZipFilePath)
    $response = ExecuteAPI -Method "Put" -Uri $PackageZipUploadUrl -Headers $Headers -Body $FileBytes -OnErrorMessage "Failed to upload action package zip!"
}

function ProcessActionPackageZip {
    param (        
        [Parameter(Mandatory)]
        [String]$PackageZipUploadUrl
    )
    
    $Uri = "$Endpoint/v1/actionPackages/processZip"

    Write-InfoLog "Processing action package zip ..."
    $Body = "{'url':'$PackageZipUploadUrl'}"   

    $ErrorMessage = "Failed to process action package zip!"
    $response = ExecuteActionAPI -Method "Post" -Uri $Uri -Body $Body -OnErrorMessage $ErrorMessage
    $MonitorPackageZipProcessingUrl = $response.url

    ExitOnError -IsError ([string]::IsNullOrWhiteSpace($MonitorPackageZipProcessingUrl)) -OnErrorMessage $ErrorMessage

    Write-DebugLog "Monitor Url: $MonitorPackageZipProcessingUrl"
    return $MonitorPackageZipProcessingUrl    
}

function MonitorStatusUrl {
    param (
        [Parameter(Mandatory = $true)]
        [String]$StatusUrl,

        [Parameter(Mandatory = $false)]
        [String]$OnErrorMessage,

        [Int32]$RetryAttempt = 0
    )
    
    if ($RetryAttempt -gt $MAX_MONITORING_RETRIES) {
        Write-ErrorLog "Max retries exhausted!"
        ExitOnError -OnErrorMessage "Max retries exhausted!"
    }
    
    $response = ExecuteActionAPI -Method "Get" -Uri $StatusUrl

    Write-DebugLog "Retry# $RetryAttempt, Status: $($response.status), SubStatus: $($response.subStatus), Message: $($response.message)"

    if ($response.status -eq "InProgress") {        
        Start-Sleep -s 2
        return MonitorStatusUrl -StatusUrl $StatusUrl -OnErrorMessage $OnErrorMessage -RetryAttempt ($RetryAttempt + 1)
    }

    $ActionPackageResourceUrl = $response.resourceUrl
    
    if ($response.status -eq "Completed" -and $response.subStatus -eq "Success") {
        Write-SuccessLog "Package processing succeeded! $($response.message)"
        return $ActionPackageResourceUrl
    }
    else {
        # Processing failed
        ExitOnError -OnErrorMessage "$OnErrorMessage, Message: $($response.message)"
    }
}

function CreateTeamsApp {
    param (
        [Parameter(Mandatory)]
        [string]$ActionPackageResourceUrl
    )
    
    $Uri = "$ActionPackageResourceUrl/teamsApp"

    Write-InfoLog "Creating Teams app ..."
    
    $ErrorMessage = "Failed to create Teams app!"

    $response = ExecuteActionAPI -Method "Post" -Uri $Uri -OnErrorMessage $ErrorMessage
    $AppCreationStatusMonitorUrl = $response.url
    ExitOnError -IsError ([string]::IsNullOrWhiteSpace($AppCreationStatusMonitorUrl)) -OnErrorMessage $ErrorMessage
    
    Write-DebugLog "Monitor App Creation status Url: $AppCreationStatusMonitorUrl"
    return $AppCreationStatusMonitorUrl    
}

function DownloadTeamsApp {
    param (
        [Parameter(Mandatory)]
        [String]$TeamsAppDownloadUrl
    )
    
    Write-InfoLog "Downloading Teams app ..."

    if ([string]::IsNullOrWhiteSpace($TeamsAppDownloadDirectoryPath) -or !(Test-Path -Path $TeamsAppDownloadDirectoryPath -PathType Container -ErrorAction Stop)) {
        $TeamsAppDownloadDirectoryPath = "$(Get-Location)\TeamsApp"
        $TeamsAppDirectory = New-Item -Path $TeamsAppDownloadDirectoryPath -Force -ItemType Directory 
    }

    $ManifestPath = "$TeamsAppDownloadDirectoryPath\TeamsManifest.zip" 

    Invoke-RestMethod -Uri $TeamsAppDownloadUrl -Method "Get" -ContentType "application/zip" -OutFile $ManifestPath   
    ExitOnError -IsError $(!(Test-Path $ManifestPath -PathType leaf)) -OnErrorMessage "Failed to download Teams app!"

    Write-SuccessLog "Teams app download succeeded! Path: $ManifestPath"    
}

#endregion ActionService APIs

#region API Invocation

$LogFilePath = Initialize-Logger $LogDirectoryPath
Write-InfoLog "RequestCorrelationId for this session: $RequestCorrelationId"

$AccessToken = ValidateOrAcquireToken

$PackageZipUploadUrl = GetActionPackageUploadUrl

UploadPackageZipToBlob $PackageZipUploadUrl $PackageZipFilePath

$MonitorPackageZipProcessingUrl = ProcessActionPackageZip $PackageZipUploadUrl

Write-InfoLog "Monitoring package processing status ..."
$ActionPackageResourceUrl = MonitorStatusUrl -StatusUrl $MonitorPackageZipProcessingUrl -OnErrorMessage "Package Processing failed! " 

$AppCreationStatusMonitorUrl = CreateTeamsApp $ActionPackageResourceUrl

Write-InfoLog "Monitoring Teams app creation status ..."
$TeamsAppDownloadUrl = MonitorStatusUrl -StatusUrl $AppCreationStatusMonitorUrl -OnErrorMessage "Teams app creation failed! "

DownloadTeamsApp $TeamsAppDownloadUrl

#endregion API Invocation