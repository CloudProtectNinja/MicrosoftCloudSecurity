<#
    .SYNOPSIS
        Starts Entra ID access reviews for specific Microsoft 365 groups and/or Microsoft Teams teams.

	.PARAMETER TenantId
		ID of the Entra ID tenant

	.PARAMETER AuthenticationType
		Type of authentication to connect to the desired Entra ID tenant
		Possible values:
			- SystemAssignedManagedIdentity
			- UserAssignedManagedIdentity
			- AppRegistration

	.PARAMETER AuthenticationUserAssignedManagedIdentityClientId
		Client ID of the user-assigned managed identity if this kind of authentication shall be used

	.PARAMETER AuthenticationAppRegistrationClientId
		Client ID of the custom app registration if this kind of authentication shall be used

	.PARAMETER AuthenticationAppRegistrationCertificateThumbprint
		Certificiate thumprint of the custom app registration if this kind of authentication shall be used

	.PARAMETER IncludedGroupsServerSideFilterQuery
		Server-side Microsoft Graph API filter query to fetch specific groups according to the following Microsoft documentation:
		https://learn.microsoft.com/en-us/graph/api/group-list

	.PARAMETER IncludedGroupsMailNicknamePrefixes
		Semicolon-separated list of mail nickname prefixes, so the access reviews are only started for these specific M365 groups and/or teams

	.PARAMETER ExcludedGroupsMailNicknamePrefixes
		Semicolon-separated list of mail nickname prefixes, so the access reviews are NOT started for these specific M365 groups and/or teams
	
	.PARAMETER IncludedGroupsVisibility
		Determines whether access reviews are only started for private or public M365 groups / teams, or for both private and public objects
		Possible values:
			- (null)
			- Private
			- Public

	.PARAMETER IncludedGroupsMinimumAgeInDays
		Determines whether access reviews are only started for M365 groups / teams that already exist in the tenant for X days

	.PARAMETER AccessReviewDurationInDays
		Number of days the group/team owners have time to answer the access review

	.PARAMETER AccessReviewDefaultDecisionIfNoOwnerResponsed
		Default decision if the owners do not answer the access review within the defined time frame. Determines whether member/guest access is removed or not if the owners do not respond.
		Possible values:
			- None
			- Approve
			- Deny
			- Recommendation

	.PARAMETER AccessReviewAdditionalMailText
		Optional text that is included in the automated e-mail sent to the owner by the Microsoft system. This string can only include plain-text and no HTML formatting.

	.PARAMETER GroupId
		If a specific object ID of a Microsoft 365 Group (or MS Teams team) is provided in this parameter, the access review is only started for this group.

    .PARAMETER DryRun
        If DryRun mode is activated, no changes will be performed in the tenant but a report about the potential changes will be generated only.

	.EXAMPLE
		# Start access reviews for private M365 groups and MS Teams teams with a "PRJ-" project prefix using managed identity-based authentication
		.\Start-AccessReviews.ps1 `
			-AuthenticationType "SystemAssignedManagedIdentity" `
			-IncludedGroupsMailNicknamePrefixes "PRJ-" `
			-IncludedGroupsVisibility "Private" `
			-AccessReviewAdditionalMailText "In case of any questions, please contact the Contoso helpdesk."

	.EXAMPLE
		# Start access reviews for private MS Teams teams, except for teams with an "DEPT-" department prefix or an "X-" prefix, using app registration-based authentication
		.\Start-AccessReviews.ps1 `
			-TenantId "aa4201e7-433f-44aa-88be-8eeb2d88a908" `
			-AuthenticationType "AppRegistration" `
			-AuthenticationAppRegistrationClientId "c091754f-c31c-4d60-ae79-53cfcd3a5d97" `
			-AuthenticationAppRegistrationCertificateThumbprint "31ED9B5036A479CE2E7180715CC232359D5E8F98" `
			-IncludedGroupsServerSideFilterQuery "groupTypes/any(c:c eq 'Unified') and resourceProvisioningOptions/Any(x:x eq 'Team')" `
			-ExcludedGroupsMailNicknamePrefixes "DEPT-;X-" `
			-IncludedGroupsVisibility "Private"

	.EXAMPLE
		# Only start a "dry run" to check, which groups/teams would be in scope of your access review, with the following command:
		.\Start-AccessReviews.ps1 `
			-AuthenticationType "SystemAssignedManagedIdentity" `
			-DryRun:$true

	.EXAMPLE
		# Only start a an access review for a single Microsoft 365 group
		.\Start-AccessReviews.ps1 `
			-AuthenticationType "SystemAssignedManagedIdentity" `
			-GroupId "3c6abeea-a4be-4633-acf6-18927400aeef"

	.NOTES
		Author: Dustin Schutzeichel (https://cloudprotect.ninja)
#>
[CmdletBinding()]
param(
	[string]
	$TenantId,
	
	[string]
	[ValidateSet("SystemAssignedManagedIdentity", "UserAssignedManagedIdentity", "AppRegistration")]
	$AuthenticationType,

	[string]
	$AuthenticationUserAssignedManagedIdentityClientId,

	[string]
	$AuthenticationAppRegistrationClientId,

	[string]
	$AuthenticationAppRegistrationCertificateThumbprint,

	[string]
	# By default, we filter for "unified" groups, i.e. Microsoft 365 Groups including Microsoft Teams teams
	# To filter for MS Teams teams only, you could use the filter "groupTypes/any(c:c eq 'Unified') and resourceProvisioningOptions/Any(x:x eq 'Team')"
	$IncludedGroupsServerSideFilterQuery = "groupTypes/any(c:c eq 'Unified')",

	[string]
	$IncludedGroupsMailNicknamePrefixes,
	
	[string]
	$ExcludedGroupsMailNicknamePrefixes,
	
	[string]
	[ValidateSet("Private", "Public")]
	# By default, we only start access reviews for private M365 groups / teams
	$IncludedGroupsVisibility = "Private",

	[int]
	$IncludedGroupsMinimumAgeInDays,

	[int]
	# By default, the owners have 30 days to answer the access review
	$AccessReviewDurationInDays = 30,

	[string]
	[ValidateSet("None", "Approve", "Deny", "Recommendation")]
	# By default, if the owners do not answer the access review, members/guests are not removed from the group/team
	$AccessReviewDefaultDecisionIfNoOwnerResponsed = "None",

	[string]
	$AccessReviewAdditionalMailText,

	[string]
	$GroupId,

    [bool]
    $DryRun
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
		if ($DryRun) {
			Write-Verbose "DryRun activated"
		}

		# Import necessary PowerShell dependencies
		Write-Verbose "Importing PowerShell modules"
		Import-PowerShellModules -ModuleNames @(
			"Microsoft.Graph.Authentication", 
			"Microsoft.Graph.Groups", 
			"Microsoft.Graph.Identity.Governance"
		)

		# Connect to Microsoft Graph API
		Write-Verbose "Connecting to Microsoft Graph API"
		Connect-ToGraph

		# Build today date object
		$today = [DateTime]::UtcNow
		$todayAtMidnightUtc = (Get-Date `
								-Year $today.Year -Month $today.Month -Date $today.Date `
								-Hour 0 -Minute 0 -Second 0 -Millisecond 0 -AsUTC).ToUniversalTime()

		Write-Verbose "--- Fetching desired groups ---"
		$groups = Get-SpecificGroupsForAccessReviews
		Write-Verbose "---"

		# Iterate all desired groups
		Write-Verbose "--- Iterating desired groups ---"
		$results = @()
		$errorCount = 0
		foreach ($group in $groups) {
			try {
				# Start access review for each group
				$result = New-AccessReview -Group $group -StartDate $todayAtMidnightUtc
				$results += $result
			} catch {
				# Catch error for specific group
				$errorCount++
				$errorMessage = Get-DetailedErrorMessage -ErrorRecord $_
				$result = [pscustomobject]@{
					"GroupId" = $group.Id
					"GroupMailNickname" = $group.MailNickname					
					"CreatedDateTime" = $group.CreatedDateTime.ToString("o")
					"Error" = $errorMessage
				}
				$results += $result
				
				Write-Warning "Unexpected error for group with id `"$($result.GroupId)`" and name `"$($result.GroupMailNickname)`":"
				Write-Warning $result.Error
			}
			
			if (!$DryRun) {				
				# Add short sleep time to mitigate Graph API throttling as good as possible
				Start-Sleep -Milliseconds 50
			}
		}
		Write-Verbose "---"

		# Output the groups with started access review and errors
		Log-Results -Results $results

		if ($errorCount -gt 0) {
			throw "Unexpected errors encountered for $errorCount groups"
		}
	} catch {
		# Log error message
		$errorMessage = Get-DetailedErrorMessage -ErrorRecord $_
		Write-Warning $errorMessage
		# Re-throw error, so runbook status is set to Failed
		throw
	} finally {
		# We do not disconnect from Graph API to avoid negative interactions with other runbooks executed in parallel
	}
}

#endregion

#region -------------------- Access Review Functions  ----------------------------------------------

Function Get-SpecificGroupsForAccessReviews {
	<#
        .SYNOPSIS
            Returns all desired M365 groups (including MS Teams teams), an access review shall be started for.
    #>
	param(
		[datetime]
		$TodayAtMidnightUtc
	)

	$additionalServerSideFilterScriptBlock = {
		param()

		$filters = @()

		# Apply server-side filter to fetch only one specific group
		if (![string]::IsNullOrEmpty($GroupId)) {			
			$groupIdFilter = "id eq '$GroupId'"
			$filters += $groupIdFilter
		}

		# TODO: Custom server-side filters can be implemented here

		$filter = $filters -join " and "
		return $filter
	}

	$additionalClientSideFilterScriptBlock = {
		param($Groups)

		# Filter for groups that are created at least X days ago		
		# Note: Server-side filter for CreatedDateTime is not possible at all due to limitations by Microsoft, so we need a client-side filter as fallback
		if ($null -ne $IncludedGroupsMinimumAgeInDays -and $IncludedGroupsMinimumAgeInDays -gt 0) {
			$minimumAgeCompareDate = $todayAtMidnightUtc.AddDays(-1 * $IncludedGroupsMinimumAgeInDays)
			$Groups = $Groups | Where-Object { $_.CreatedDateTime -le $minimumAgeCompareDate }
			Write-Verbose "$($Groups.Count) groups after filtering for minimum group age"
		}

		# Apply client-side filter to fetch only one specific group
		if (![string]::IsNullOrEmpty($GroupId)) {
			$Groups = $Groups | Where-Object { $_.Id -eq $GroupId }
			Write-Verbose "$($Groups.Count) groups after filtering for one specific group"
		}

		# TODO: Custom client-side filters can be implemented here

		return $Groups
	}

	$groups = Get-Groups	`
		-SelectProperties @(
			"id", "displayName", "mailNickname", "createdDateTime") `
		-OrderByProperty "displayName" `
		-IsAdvancedQuery:$true `
		-AdditionalServerSideFilterScriptBlock $additionalServerSideFilterScriptBlock `
		-AdditionalClientSideFilterScriptBlock $additionalClientSideFilterScriptBlock

	return $groups
}

Function Get-Groups {
	<#
        .SYNOPSIS
            Returns all Microsoft 365 groups, which are in scope of the governance processes.
    #>
	param(
		[string[]]
		$SelectProperties,

		[string[]]
		$ExpandProperties,

		[string]
		$OrderByProperty,

		[boolean]
		$IsAdvancedQuery,

		[scriptblock]
		$AdditionalServerSideFilterScriptBlock,

		[scriptblock]
		$AdditionalClientSideFilterScriptBlock
	)

	### 1. Prepare the $select properties ###

	# Adjust the $select properties if necessary
	if ($null -eq $SelectProperties) {
		$SelectProperties = @()
	}
	
	# Always add certain properties to the $select statement because we probably need it to perform client-side filtering and further processing
	$SelectProperties += "id"
	$SelectProperties += "visibility"
	$SelectProperties += "mailNickname"
	$SelectProperties += "displayName"
	$SelectProperties += "membershipRule"
	$SelectProperties += "isAssignableToRole"

	# Remove duplicates from selected properties
	$SelectProperties = $SelectProperties | Sort-Object -Unique
	Write-Verbose "`$select=$($SelectProperties -join ",")"

	###################################

	### 2. Prepare the $expand property ###	

	$hasExpandProperties = $null -ne $ExpandProperties -and $ExpandProperties.Count -gt 0	

	if ($hasExpandProperties) {
		Write-Verbose "`$expand=$($ExpandProperties -join ",")"	
	}

	#######################################

	### 3. Prepare the $filter property ###

	# Build server-side $filter for groups/teams in scope
	$filters = @()

	# Add a common filter query if necessary (typically: groupTypes/any(c:c eq 'Unified'))
	if (![string]::IsNullOrEmpty($IncludedGroupsServerSideFilterQuery)) {
		$filters += $IncludedGroupsServerSideFilterQuery
	}

	# Filter for groups with specific prefixes
	if (![string]::IsNullOrEmpty($IncludedGroupsMailNicknamePrefixes)) {
		$prefixes = $IncludedGroupsMailNicknamePrefixes -split ";"
		$prefixFilters = @()
		foreach ($prefix in $prefixes) {
			$prefixFilters += "startsWith(mailNickname, '$prefix')"
		}
		$prefixFilter = $prefixFilters -join " or "
		$filters += "($prefixFilter)" # Due to the "or" query with have to encapsulate the prefix filter in paranthesis
	}

	# We can only apply the following "advanced query" filters at server-side if we do not have an $expand operation at the same time as well
	# according to the limitations documented under:
	# https://learn.microsoft.com/en-us/graph/aad-advanced-queries?tabs=http#query-scenarios-that-require-advanced-query-capabilities
	if (!$hasExpandProperties) {
		# Filter for groups without specific prefixes
		if (![string]::IsNullOrEmpty($ExcludedGroupsMailNicknamePrefixes)) {
			$prefixes = $ExcludedGroupsMailNicknamePrefixes -split ";"
			$prefixFilters = @()
			foreach ($prefix in $prefixes) {
				$prefixFilters += "not(startsWith(mailNickname, '$prefix'))"
			}
			$prefixFilter = $prefixFilters -join " and "
			
			$filters += $prefixFilter
			$IsAdvancedQuery = $true # Due to the "not()" operator, this is an advanced query
		}
	}

	# Apply script block for deliverable-specific filter logic
	if ($null -ne $AdditionalServerSideFilterScriptBlock) {
		$additionalFilter = Invoke-Command -ScriptBlock $AdditionalServerSideFilterScriptBlock
		if (![string]::IsNullOrEmpty($additionalFilter)) {
			$filters += $additionalFilter
		}
	}

	# Combine the server-side filter for groups
	$filter = $filters -join " and "
	Write-Verbose "`$filter=$filter"

	#######################################    

    ### 4. Prepare the $orderby properties ###	
    $orderby = $null

    if (!$hasExpandProperties -and ![string]::IsNullOrEmpty($OrderByProperty)) {
		# We can apply $orderby for server-side sorting only if no $expand is applied
		$orderby = $OrderByProperty
		Write-Verbose "`$orderby=$orderby"

        if (![string]::IsNullOrEmpty($filter)) {
            # $orderby and $filter at the same time is only allowed in the advanced query mode:
            # https://learn.microsoft.com/en-us/graph/aad-advanced-queries?tabs=http#query-scenarios-that-require-advanced-query-capabilities
            $IsAdvancedQuery = $true
        }		
	}

    #######################################

	### 5. Perform the Microsoft Graph API call ###

	# Build the parameters for the Get-MgGroup cmdlet
	$groupParams = @{
		"Property" = $SelectProperties
		"All" = $true
		"ErrorAction" = "Stop"
		"Verbose" = $false
	}

    # Add $expand parameter
    if ($hasExpandProperties) {
        $groupParams.Add("ExpandProperty", $ExpandProperties)
    }

    # Add $filter parameter
    if (![string]::IsNullOrEmpty($filter)) {
        $groupParams.Add("Filter", $filter)
    }

    # Add $orderby parameter
    if (![string]::IsNullOrEmpty($orderby)) {
        $groupParams.Add("Sort", $orderby)
    }

	# In case of an advanced query, add the ConsistencyLevel and CountVariable parameters
	if ($IsAdvancedQuery) {
		$groupParams.Add("ConsistencyLevel", "eventual")
		$groupParams.Add("CountVariable", "groupsCount")
		Write-Verbose "`$count=true"
	}

    # Fetch all relevant groups from the server
	$groups = Get-MgGroup @groupParams
	Write-Verbose "$($groups.Count) groups returned by server-side filter"

	#######################################

	### 6. Apply client-side filters and sorting if necessary ###

	# Perform client-side filtering to exclude role-assignable groups
	$groups = $groups | Where-Object { $null -eq $_.IsAssignableToRole -or $_.IsAssignableToRole -eq $false }
	Write-Verbose "$($groups.Count) groups after excluding role-assignable groups"

	# Perform client-side filtering to exclude dynamic groups
	$groups = $groups | Where-Object { $null -eq $_.MembershipRule }
	Write-Verbose "$($groups.Count) groups after excluding dynamic groups"

	# Perform client-side filtering for prefixes
	if (![string]::IsNullOrEmpty($IncludedGroupsMailNicknamePrefixes)) {
		$prefixes = $IncludedGroupsMailNicknamePrefixes -split ";"
		$prefixMatch = "^(" + ($prefixes -join "|") + ")"
		$groups = $groups | Where-Object { $_.MailNickname -match $prefixMatch }
		Write-Verbose "$($groups.Count) groups after filtering for inclusion prefixes"
	}

	if (![string]::IsNullOrEmpty($ExcludedGroupsMailNicknamePrefixes)) {
		$prefixes = $ExcludedGroupsMailNicknamePrefixes -split ";"
		$prefixMatch = "^(" + ($prefixes -join "|") + ")"
		$groups = $groups | Where-Object { $_.MailNickname -notmatch $prefixMatch }
		Write-Verbose "$($groups.Count) groups after filtering for exclusion prefixes"
	}
	
	# Perform client-side filtering for group visibility (private/public)
	if (![string]::IsNullOrEmpty($IncludedGroupsVisibility)) {
		$groups = $groups | Where-Object { $_.Visibility -eq $IncludedGroupsVisibility }
		Write-Verbose "$($groups.Count) groups after filtering for group visibility"
	}

	# Apply script block for deliverable-specific filter logic
	if ($null -ne $AdditionalClientSideFilterScriptBlock) {
		$groups = Invoke-Command -ScriptBlock $AdditionalClientSideFilterScriptBlock -ArgumentList (,$groups)
	}

	# Perform client-side sorting
	if ([string]::IsNullOrEmpty($orderby) -and ![string]::IsNullOrEmpty($OrderByProperty)) {
		$groups = $groups | Sort-Object -Property $OrderByProperty
		Write-Verbose "Sorted by $OrderByProperty"
	}

	#######################################	

	return $groups
}

Function New-AccessReview {
	<#
        .SYNOPSIS
            Starts a new access review for a specific group.
    #>
	param(
		$Group,

		[datetime]
		$StartDate
	)

	# Build start and end date variables for the new access review
	$startDateString = $StartDate.ToString("yyyy-MM-dd")
	$endDate = $StartDate.AddDays($AccessReviewDurationInDays + 1)
	$endDateString = $endDate.ToString("yyyy-MM-dd")

	# Configure whether members/guests shall be removed if the owner(s) do not respond to the access review request
	$defaultDecisionEnabled = $false
	$defaultDecision = "None"
	if (![string]::IsNullOrEmpty($AccessReviewDefaultDecisionIfNoOwnerResponsed)) {
		$defaultDecision = $AccessReviewDefaultDecisionIfNoOwnerResponsed
		if ($defaultDecision -ne "None") {
			$defaultDecisionEnabled = $true
		}
	}

	# Define all access review parameters
	$accessReviewParams = @{
		# Name for the group-specific access review
		displayName = $Group.MailNickname

		# Additional text paragraph displayed in the access review mail sent to the owners (text only, HTML not supported)
		descriptionForReviewers = $AccessReviewAdditionalMailText
	
		# All members & guests including B2B Direct Connect users will be reviewed by the group owners
		scope = @{
			"@odata.type" = "#microsoft.graph.principalResourceMembershipsScope"
			principalScopes = @(
				@{
					"@odata.type" = "#microsoft.graph.accessReviewQueryScope"
					query = "/users"
					queryType = "MicrosoftGraph"
					queryRoot = $null
				}
			)
			resourceScopes = @(
				# Internal members and guests have to be reviewed
				@{
					"@odata.type" = "#microsoft.graph.accessReviewQueryScope"
					query = "/groups/$($Group.Id)/transitiveMembers"
					queryType = "MicrosoftGraph"
					queryRoot = $null
				},
				# B2B direct connect users of shared channels in MS Teams have to be reviewed as well
				@{
					"@odata.type" = "#microsoft.graph.accessReviewQueryScope"
					query = "/teams/$($Group.Id)/channels?`$filter=(membershipType eq 'shared')"
					queryType = "MicrosoftGraph"
					queryRoot = $null
				}
			)
		}
	
		# Owners are responsible for the review of their group
		reviewers = @(
			@{
				query = "/groups/$($Group.Id)/owners"
				queryType = "MicrosoftGraph"
			}
		)
	
		settings = @{
			# Owners have X days to complete the access review
			instanceDurationInDays = $AccessReviewDurationInDays
	
			# Microsoft's default e-mail notifications + reminders are sent to the owners
			mailNotificationsEnabled = $true
			reminderNotificationsEnabled = $true
	
			# Owners do not need to write a justification text in order to extend/remove access for member/guest XY
			justificationRequiredOnApproval = $false
	
			# Show recommendations to owners whether member/guest XY shall be extended/removed based on the last sign-in activity
			# https://learn.microsoft.com/en-us/graph/api/resources/accessreviewschedulesettings?view=graph-rest-1.0
			# https://learn.microsoft.com/en-us/azure/active-directory/governance/licensing-fundamentals#features-by-license
			recommendationsEnabled = $true
			recommendationLookBackDuration = "P90D" # see https://learn.microsoft.com/en-us/graph/api/resources/accessreviewschedulesettings?view=graph-rest-1.0
			
			# Configuration whether members/guests shall be removed if the owner(s) do not respond to the access review request
			defaultDecisionEnabled = $defaultDecisionEnabled
			defaultDecision = $defaultDecision
			autoApplyDecisionsEnabled = $true
			applyActions = @()
	
			# No recurrence: the access review for the group is only performed once
			# To periodically review access for the groups, you can create an Azure Automation runbook schedule and execute the runbook every X months, for example
			recurrence = @{
				pattern = $null
				range = @{
					type = "endDate"
					numberOfOccurrences = 0
					recurrenceTimeZone = 0
					# Set start and end date for the access review
					startDate = $startDateString
					endDate = $endDateString
				}
			}
		}

		# No recurrence
		instanceEnumerationScope = $null

		# No backup reviews and no other mail recipients at the end of the review
		backupReviewers = @()
   		additionalNotificationRecipients = @()
	}

	if (!$DryRun) {
		# Start the access review
		New-MgIdentityGovernanceAccessReviewDefinition -BodyParameter $accessReviewParams -ErrorAction Stop -Verbose:$false | Out-Null
	}

	# Return the results for this group
	return [pscustomobject]@{
		"GroupId" = $Group.Id
		"GroupMailNickname" = $Group.MailNickname
		"CreatedDateTime" = $Group.CreatedDateTime.ToString("o")
	}
}

Function Log-Results {
	<#
        .SYNOPSIS
            Write the results to the output stream.
    #>
	param(
		$Results
	)

	Write-Verbose "--- Output groups with started access review ---"
	$groupsWithStartedAccessReview = $Results | Where-Object { [string]::IsNullOrEmpty($_.Error) }
	if ($groupsWithStartedAccessReview.Count -gt 0) {
		# Output groups in batches to mitigate the Azure runbook output limitations
		$batchSize = 1000
		for ($i = 0; $i -lt $groupsWithStartedAccessReview.Count; $i += $batchSize) {
			$outputCsv = $null
			if (($groupsWithStartedAccessReview.Count - $i) -ge $batchSize) {
				$outputCsv = $groupsWithStartedAccessReview[$i..($i + $batchSize - 1)] | ConvertTo-Csv -Delimiter ";" -NoTypeInformation
			} elseif ($groupsWithStartedAccessReview.Count -eq 1) {
				$outputCsv = $groupsWithStartedAccessReview | ConvertTo-Csv -Delimiter ";" -NoTypeInformation
		 	} else {
				$outputCsv = $groupsWithStartedAccessReview[$i..($groupsWithStartedAccessReview.Count - 1)] | ConvertTo-Csv -Delimiter ";" -NoTypeInformation
			}

			$outputCsvString = $outputCsv -join "`n"
			Write-Verbose $outputCsvString				
		}
	}
	Write-Verbose "---"	

	Write-Verbose "--- Output groups with errors ---"
	$groupsWithErrors = $Results | Where-Object { ![string]::IsNullOrEmpty($_.Error) }
	if ($groupsWithErrors.Count -gt 0) {
		# Output groups in batches to mitigate the Azure runbook output limitations
		$batchSize = 10
		for ($i = 0; $i -lt $groupsWithErrors.Count; $i += $batchSize) {
			$outputCsv = $null
			if (($groupsWithErrors.Count - $i) -ge $batchSize) {
				$outputCsv = $groupsWithErrors[$i..($i + $batchSize - 1)] | ConvertTo-Csv -Delimiter ";" -NoTypeInformation
			} elseif ($groupsWithErrors.Count -eq 1) {
				$outputCsv = $groupsWithErrors | ConvertTo-Csv -Delimiter ";" -NoTypeInformation
		 	} else {
				$outputCsv = $groupsWithErrors[$i..($groupsWithErrors.Count - 1)] | ConvertTo-Csv -Delimiter ";" -NoTypeInformation
			}

            $outputCsvString = $outputCsv -join "`n"
            Write-Verbose $outputCsvString			
		}
	}
	Write-Verbose "---"
	
}

#endregion

#region -------------------- Helper Functions: PowerShell Module Dependencies ------------------------------

Function Import-PowerShellModules {
    <#
        .SYNOPSIS
            Resolves the script dependencies by importing ncessary PowerShell modules.
    #>
    param(
		[string[]]
		$ModuleNames
	)

	if ($null -eq $ModuleNames) {
		return
	}

    # Temporarily disable the verbose logging for module import
    $previousVerbosePreference = $VerbosePreference
    $script:VerbosePreference = "SilentlyContinue"
    
	foreach ($moduleName in $ModuleNames) {
		Import-Module -Name $moduleName -Verbose:$false
	}

    # Reset the verbose logging
    $script:VerbosePreference = $previousVerbosePreference
}

#endregion

#region -------------------- Helper Functions: Connection/Authentication Functions -------------------------

Function Connect-ToGraph {
    <#
        .SYNOPSIS
            Connect to Microsoft Graph API.
    #>
    param()

	if ($AuthenticationType -eq "SystemAssignedManagedIdentity") {
		# Connect via managed identity
		Connect-MgGraph `
			-Identity `
			-ErrorAction Stop -Verbose:$false | Out-Null
	} elseif ($AuthenticationType -eq "UserAssignedManagedIdentity") {
		# Connect via managed identity
		Connect-MgGraph `
			-Identity `
			-ClientId $AuthenticationUserAssignedManagedIdentityClientId `
			-ErrorAction Stop -Verbose:$false | Out-Null
	} elseif ($AuthenticationType -eq "AppRegistration") {
		# Connect with the certificate uploaded into the Azure Automation certificate store via thumbprint
		Connect-MgGraph `
			-TenantId $TenantId `
			-ClientId $AuthenticationAppRegistrationClientId `
			-CertificateThumbprint $AuthenticationAppRegistrationCertificateThumbprint `
			-ErrorAction Stop -Verbose:$false | Out-Null		
	} else {
		throw "Unsupported authentication type for Connect-MgGraph command"
	}

	# Increase retries in case of Microsoft Graph API related issues
	# https://github.com/microsoftgraph/msgraph-sdk-powershell/pull/1584#issue-1417492648
	Set-MgRequestContext -MaxRetry 10 -ErrorAction Stop -Verbose:$false | Out-Null
}

#endregion

#region -------------------- Helper Functions: Logging/Error Handling -------------------

Function Get-DetailedErrorMessage {
    <#
        .SYNOPSIS
            Get details of the error message as string.
    #>  
    param(
        [System.Management.Automation.ErrorRecord]
        $ErrorRecord
    )

    $errorDetails = $ErrorRecord | Out-String
	return $errorDetails
}

#endregion

# Run Main
Main
