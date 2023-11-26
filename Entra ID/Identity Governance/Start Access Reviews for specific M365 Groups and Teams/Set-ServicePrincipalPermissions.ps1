<#
    .SYNOPSIS
        Assigns appropriate Microsoft Graph API permissions to the stated service principal (managed identity or custom app registration).

    .PARAMETER ServicePrincipalObjectId
        Object ID of the service principal (managed identity / app) you want to assign the API permission to
    
	.EXAMPLE
		.\Set-ServicePrincipalPermissions.ps1 -ServicePrincipalObjectId "c091754f-c31c-4d60-ae79-53cfcd3a5d97"
    
    .NOTES  
        Author: Dustin Schutzeichel (https://cloudprotect.ninja)
#>
[CmdletBinding()]
param(
    [string]
    $ServicePrincipalObjectId
)

$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"

# Connect to Microsoft Graph API with the necessary permission scopes
Connect-MgGraph `
    -Scopes Application.Read.All, AppRoleAssignment.ReadWrite.All
    -NoWelcome `
    -ErrorAction Stop -Verbose:$false

# Configure Microsoft Graph API permissions for the service principal
$graphApiPermissions = @(
    "Group.Read.All", # Needed to read attributes of M365 group / MS Teams team objects
    "AccessReview.ReadWrite.All" # Needed to initiate new access reviews
)

$graphResource = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'" -ErrorAction Stop -Verbose:$false
foreach ($permission in $graphApiPermissions) {
    $role = $graphResource.AppRoles | Where-Object { $_.Value -eq $permission }
    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $ServicePrincipalObjectId `
        -PrincipalId $ServicePrincipalObjectId `
        -AppRoleId $role.Id `
        -ResourceId $graphResource.Id `
        -ErrorAction Stop -Verbose:$false
}
