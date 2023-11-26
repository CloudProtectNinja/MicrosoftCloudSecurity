<#
    .SYNOPSIS
        This script creates an export of all "SharePoint-only" guest users in comparison with Microsoft Entra B2B guest users 
        across all SharePoint and/or OneDrive sites of your Microsoft 365 tenant.

    .DESCRIPTION
        Required PowerShell modules:
        - PnP.PowerShell: https://github.com/pnp/powershell / Script developed with version 1.12.0

        Required API permissions for the registered Entra ID application:
        - Sites.FullControl.All application permissions on the SharePoint API (https://microsoft.sharepoint-df.com/Sites.FullControl.All)
    
    .PARAMETER TenantName
        The name of the Microsoft 365 tenant, i.e. https://[TenantName].sharepoint.com.
    
    .PARAMETER ClientId
        The client id of the Entra ID application, which has "Sites.FullControl.All" SharePoint application permissions to access all SharePoint and OneDrive sites.     

    .PARAMETER CertificateThumbprint
        The thumbrint of the certificate associated with the Entra ID application, which is used to find the certificate in the Windows Certificate Store.

    .PARAMETER IncludedServices
        Defines whether to include SharePoint and OneDrive sites or only of the services.

    .PARAMETER SiteUrlInputFile
        In case you want to iterate only over specific SharePoint/OneDrive sites, you can create a CSV file with a "SiteUrl" column, to be used as input for the script.

    .PARAMETER LocalOutputPath
        In case that this script is run locally, an output directory for log files must be provided.
        By default, an "Output" directory in the current script folder is created automatically.
        
    .PARAMETER CsvDelimiter
        The delimiter used in CSV files to separate columns from each other.
        
    .EXAMPLE
        # Create export on guest users across SharePoint and OneDrive
        .\Export-SharePointOnlyGuestUsers.ps1 -TenantName "yourtenant" -ClientId "fa1a81f1-e729-44d8-bb71-0a0c339c0f62" -CertificateThumbprint "789123ABC072E4125785GH4F836AFB12FA64DB210"
    
    .EXAMPLE
        # Create export on guest users across OneDrive
        .\Export-SharePointOnlyGuestUsers.ps1 -IncludedServices @("OneDrive") -TenantName "yourtenant" -ClientId "fa1a81f1-e729-44d8-bb71-0a0c339c0f62" -CertificateThumbprint "789123ABC072E4125785GH4F836AFB12FA64DB210"

    .EXAMPLE
        # Create export on guest users across the sites specified in the provided CSV file (with "SiteUrl" column)
        .\Export-SharePointOnlyGuestUsers.ps1 -SiteUrlInputFile ".\Sites.csv" -TenantName "yourtenant" -ClientId "fa1a81f1-e729-44d8-bb71-0a0c339c0f62" -CertificateThumbprint "789123ABC072E4125785GH4F836AFB12FA64DB210"

    .LINK
        https://cloudprotect.ninja/reporting-sharepoint-only-guest-users/

    .LINK
        https://github.com/CloudProtectNinja/MicrosoftCloudSecurity/tree/main/Azure%20AD/B2B%20Collaboration/Export%20SharePoint-only%20Guest%20Users

    .NOTES
        Author: Dustin Schutzeichel (https://cloudprotect.ninja)
#>

[CmdletBinding()]
param(
    [string]
    [Parameter(Mandatory = $true)]
    $TenantName,

    [string]
    [Parameter(Mandatory = $true)]
    $ClientId,
    
    [string]
    [Parameter(Mandatory = $true)]
    $CertificateThumbprint,

    [string[]]
    [Parameter(Mandatory = $false)]
    $IncludedServices = @("SharePoint", "OneDrive"),

    [string]
    [Parameter(Mandatory = $false)]
    $SiteUrlInputFile,

    [string]
    [Parameter(Mandatory = $false)]
    $LocalOutputPath = ".\Output\",

    [string]
    [Parameter(Mandatory = $false)]
    $CsvDelimiter = ";"
)

$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"

#region -------------------- Main Function  ----------------------------------------------

Function Main {
    <#
        .SYNOPSIS
            Defines the main logic of the script.
    #>
    
    try {
        # Start logging
        Start-Logging

        # Import PnP PowerShell module to have the typings available
        Import-Module -Name "PnP.PowerShell" -Verbose:$false

        # Export the external users from all SharePoint sites
        Export-GuestUsers
    } catch {
        # Log fatal script error
        Write-ScriptError -ErrorRecord $_
        # Rethrow error
        throw
    } finally {
        # Disconnect from Microsoft 365 / Entra ID tenant
        Disconnect

        # Stop logging
        Stop-Logging
    }
}

#endregion

#region -------------------- Connection/Authentication Functions -------------------------

Function Connect-ToPnPSharePoint {
    <#
        .SYNOPSIS
            Connect to SharePoint Online via the PnP module.
    #>
    [OutputType([PnP.PowerShell.Commands.Base.PnPConnection])]
    param(
        [switch]
        [Parameter(Mandatory = $false, ParameterSetName = "TenantAdministration")]
        $TenantAdministration,

        [string]
        [Parameter(Mandatory = $true, ParameterSetName = "SiteUrl")]
        $SiteUrl
    )    
    
    # Build SharePoint URL to connect with
    $url = $null
    if ($TenantAdministration) {
        # Connect to SharePoint Online tenant administration
        $url = "https://$TenantName-admin.sharepoint.com"          
    } else {
        # Connect to specific site
        $url = $SiteUrl
    }

    # Connect to SharePoint Online via PnP
    $connection = $null
    if (![string]::IsNullOrEmpty($ClientId) -and ![string]::IsNullOrEmpty($CertificateThumbprint)) {
        # Connect via Entra ID application credentials (client id + certificate)
        
        # see https://github.com/pnp/powershell/blob/dev/samples/Connect.AzureAutomation/test-connection-runbook.ps1
        # see https://github.com/pnp/powershell/tree/dev/samples/Connect.AzureAutomation

        # Build the tenant .onmicrosoft.com address
        $tenant = "$TenantName.onmicrosoft.com"

        # Connect via OAuth 2.0 client credentials flow with Entra ID app identity
        # Find certificate from Windows Certificate Store by thumbprint
        $connection = Connect-PnPOnline -Url $url `
            -ClientId $ClientId `
            -Thumbprint $CertificateThumbprint `
            -Tenant $tenant `
            -ReturnConnection `
            -Verbose:$false       
    } else {
        throw "Please provide the ClientId and CertificateThumbprint parameters for app-only authentication"
    }

    if ($null -ne $connection) {
        # Return the PnP connection object
        return $connection
    }

    # Otherwise the connection has failed
    throw "PnP connection failed"
}

Function Disconnect {
    <#
        .SYNOPSIS
            Closes all connections.
    #>

    try {
        # Disconnect all connections from SharePoint Online
        Disconnect-PnPOnline
    } catch {
        # Errors may occur here when already disconnected, but can be ignored
    }
}

#endregion

#region -------------------- External User Functions -------------------------------------

Function Export-GuestUsers {
    <#
        .SYNOPSIS
            Exports all external users from SharePoint Online / OneDrive for Business sites into a CSV file.
    #>
    param()

    $logDate = Get-Date -Format "yyyy-MM-dd-HH-mm"

    # Connect to SharePoint tenant administration
    $PnPTenantConnection = $null
    try {      
        $PnPTenantConnection = Connect-ToPnPSharePoint -TenantAdministration
        Write-Host "Connected to SharePoint tenant administration"
    } catch {
        Write-ErrorLogEntry -ErrorRecord $_ -Message "Failed to connect to SharePoint tenant administration"
        return
    }

    # Get SharePoint sites to iterate over
    $sites = @()
    try {
        if (![string]::IsNullOrEmpty($SiteUrlInputFile)) {
            # Get the site URLs from the CSV
            $sitesFromInputFile = Import-Csv -LiteralPath $SiteUrlInputFile -Delimiter $CsvDelimiter -Encoding UTF8 | Select-Object -Property "SiteUrl"
            foreach ($site in $sitesFromInputFile) {
                $props = @{
                    Url = $site.SiteUrl
                }
                $sites += $props
            }
        } else {
            # Please note, due to a bug in SharePoint CSOM we cannot provide a server-side filter to only return sites with enabled external sharing state.
            # Consequently, we need to fetch ALL sites from the server.
            # For more details and the current status of the reported issue, please see:
            # https://github.com/pnp/powershell/issues/2442

            if ($null -eq $IncludedServices -or ($IncludedServices -contains "SharePoint" -and $IncludedServices -contains "OneDrive")) {
                # Include both SharePoint and OneDrive sites
                $sites = Get-PnPTenantSite -IncludeOneDriveSites -Connection $PnPTenantConnection
            } elseif ($IncludedServices -contains "SharePoint") {
                # Include only SharePoint sites
                $sites = Get-PnPTenantSite -Connection $PnPTenantConnection
            } elseif ($IncludedServices -contains "OneDrive") {
                $sites = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like 'https://$TenantName-my.sharepoint.com/personal/'" -Connection $PnPTenantConnection
            } else {
                throw "Unexpected value in IncludedServices script parameter"
            }
        }
    } catch {
        Write-ErrorLogEntry -ErrorRecord $_ -Message "Failed to fetch tenant sites" 
        return
    }

    foreach ($site in $sites) {
        # Process current site
        $siteUrl = $site.Url
        Write-Host
        Write-Host $siteUrl

        # Fetch and export details about external users in the current site
        try {          
            # Open connection to current SharePoint site
            $PnPSiteConnection = Connect-ToPnPSharePoint -SiteUrl $siteUrl

            # Get site title if we read the CSV input
            $siteTitle = $site.Title
            if ([string]::IsNullOrEmpty($siteTitle)) {
                $web = Get-PnPWeb -Connection $PnPSiteConnection
                $siteTitle = $web.Title
            }

            # Get external users based on the "urn:spo:guest" or "#EXT#" identifier  
            $externalUsers = Get-PnPUser -Connection $PnPSiteConnection | Where-Object { 
                $_.LoginName -like "*urn%3aspo%3aguest*" -or $_.LoginName -like "*urn:spo:guest*" -or $_.LoginName -like "*#EXT#*" 
            }            

            foreach ($user in $externalUsers) {
                # Differentiate between SPO-only guest vs. Entra B2B guest
                $externalUserType = "SPO"
                if ($user.LoginName -like "*#EXT#*") {
                    $externalUserType = "B2B"
                }

                # Specify the properties of the external user to be exported
                $props = [ordered]@{
                    SiteUrl                         = $siteUrl
                    SiteTitle                       = $site.Title
                    ExternalUserType                = $externalUserType
                    Email                           = $user.Email                    
                    EmailWithFallback               = $user.EmailWithFallback
                    LoginName                       = $user.LoginName
                    Title                           = $user.Title
                    Id                              = $user.Id
                    UserId                          = $user.UserId
                    UserPrincipalName               = $user.UserPrincipalName
                    AadObjectId                     = $user.AadObjectId
                    IsShareByEmailGuestUser         = $user.IsShareByEmailGuestUser
                    IsEmailAuthenticationGuestUser  = $user.IsEmailAuthenticationGuestUser
                    IsHiddenInUI                    = $user.IsHiddenInUI
                    IsSiteAdmin                     = $user.IsSiteAdmin
                    Groups                          = ($user.Groups.Id -join ", ")
                    Expiration                      = $user.Expiration
                }
            
                # Append to CSV
                $row = New-Object -TypeName PSObject -Property $props
                $row | Export-Csv -LiteralPath "$($LocalOutputPath)\GuestUsers_$($logDate).csv" -Delimiter $CsvDelimiter -Encoding UTF8 -NoTypeInformation -Append
            }

            Write-Host "Exported $($externalUsers.Count) external users"
        } catch {
            Write-ErrorLogEntry -ErrorRecord $_ -Message "Failed to fetch the external users of the current SharePoint site"

            # Write site with error to CSV
            $errorProps = [ordered]@{
                SiteUrl                         = $siteUrl
                ErrorRecord                     = $_
                Exception                       = $_.Exception
                ScriptStackTrace                = $_.ScriptStackTrace
            }

            $row = New-Object -TypeName PSObject -Property $errorProps
            $row | Export-Csv -LiteralPath "$($LocalOutputPath)\Errors_$($logDate).csv" -Delimiter $CsvDelimiter -Encoding UTF8 -NoTypeInformation -Append
        }
    }
}

#endregion

#region -------------------- Helper Functions: Logging/Error Handling -------------------

Function Write-ErrorLogEntry {
    <#
        .SYNOPSIS
            Write a new entry to the error log.
    #>  
    param(
        [string]
        [Parameter(Mandatory = $true)]
        $Message,

        [System.Management.Automation.ErrorRecord]
        [Parameter(Mandatory = $false)]
        $ErrorRecord
    )

    # Write to PowerShell warning stream
    if ($null -ne $Message) {
        $currentDate = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
        Write-Warning "[$currentDate] ERROR: $Message"
    }

    if ($null -ne $ErrorRecord) {
        Write-Callstack -ErrorRecord $ErrorRecord
    }
}

Function Start-Logging {
    <#
        .SYNOPSIS
            Initialized the logging and starts the PowerShell transcript for local logging.
    #>    
       
    # Create output folder first
    $outputFolders = @($LocalOutputPath)
    foreach ($folder in $outputFolders) {
        if (-not (Test-Path -LiteralPath $folder)) {
            New-Item $folder -ItemType Directory | Out-Null
        }
    }

    # Start transcript
    $transcriptLogDate = Get-Date -Format "yyyy-MM-dd-HH-mm"
    Start-Transcript -LiteralPath "$($LocalOutputPath)\_Transcript_$($transcriptLogDate).log"
}

Function Stop-Logging {
    <#
        .SYNOPSIS
            Stops the logging.
    #>  
   
    try {
        # Stop the local transcript log
        Stop-Transcript
    } catch {
        Write-Warning "Failed to stop the transcript: $($_.Exception.Message)"
    }    
}

Function Write-ScriptError {
    <#
        .SYNOPSIS
            Logs a fatal script error.
    #>  
    param(
        [System.Management.Automation.ErrorRecord]
        [Parameter(Mandatory = $true)]
        $ErrorRecord,
        
        [string]
        [Parameter(Mandatory = $false)]
        $Message
    )

    # Write warning message
    if ($null -ne $Message) {
        Write-Warning $Message
    }

    # Write out the whole callstack if available
    Write-Callstack -ErrorRecord $ErrorRecord
}

Function Write-Callstack {
    <#
        .SYNOPSIS
            Logs the callstack associated with an error.
    #>  
    param(
        [System.Management.Automation.ErrorRecord]
        [Parameter(Mandatory = $true)]
        $ErrorRecord
    )

    if ($ErrorRecord) {
        Write-Warning "$ErrorRecord $($ErrorRecord.InvocationInfo.PositionMessage)"

        # Log the exception
        if ($ErrorRecord.Exception) {
            Write-Warning $ErrorRecord.Exception
        }

        if ($null -ne (Get-Member -InputObject $ErrorRecord -Name ScriptStackTrace)) {
            # PowerShell 3.0 has a stack trace on the ErrorRecord
            Write-Warning $ErrorRecord.ScriptStackTrace
        }
    }
}

#endregion

# Run Main function
Main
