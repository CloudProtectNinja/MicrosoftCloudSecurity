<a name="readme-top"></a>



<!-- PROJECT SHIELDS -->
[![Contributors][contributors-shield]][contributors-url]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![MIT License][license-shield]][license-url]



<!-- TITLE -->
# Starting Access Reviews for specific Microsoft 365 Groups and Teams using Azure Runbooks



<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li><a href="#overview">Overview</a></li>
    <li><a href="#prerequisites">Prerequisites</a></li>
    <li><a href="#usage">Usage</a></li>
    <li><a href="#contributing">Contributing</a></li>
    <li><a href="#contact">Contact</a></li>
  </ol>
</details>



<!-- OVERVIEW -->
## Overview

The Azure runbook [Start-AccessReviews.ps1](Start-AccessReviews.ps1) allows to start [Entra ID access reviews](https://learn.microsoft.com/en-us/entra/id-governance/access-reviews-overview) for specific Microsoft 365 Groups and/or Microsoft Teams teams. To be more precise, you can tailor the access reviews as follows:
- only include Microsoft 365 groups or only Microsoft Teams teams, or both
- only include groups/teams with a certain prefix
- exclude groups/teams with a certain prefix
- only include groups/teams with private or public visibility, or both
- only include groups/teams that already exist in the tenant for X days
- only start the access review for one specific group/team if desired
- combine all these options with each other according to your needs
- easily implement your own custom filter to include/exclude certain groups/teams based on other group attributes

Please find more details in my blog post:

[Starting Access Reviews for specific M365 Groups and Teams using Azure Runbooks - Cloud Protect Ninja][blog-post-url]

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- PREREQUISITES -->
## Prerequisites

In order to execute the Azure runbook, the following prerequisite steps must be performed:

1. Meet the Microsoft licensing requirements (e.g. enough Entra ID P2 licenses) according to the [Microsoft documentation](https://learn.microsoft.com/en-us/entra/id-governance/licensing-fundamentals).

2. Create an Azure Automation account according to the [Microsoft documentation](https://learn.microsoft.com/en-us/azure/automation/quickstarts/create-azure-automation-account-portal).

3. Import the following PowerShell modules with runtime version `7.2 (preview)` into the Azure Automation account using the `Browse gallery` option according to the [Microsoft documentation](https://learn.microsoft.com/en-us/azure/automation/automation-runbook-gallery):
   - [Microsoft.Graph.Authentication](https://www.powershellgallery.com/packages/Microsoft.Graph.Authentication/)
   - [Microsoft.Graph.Groups](https://www.powershellgallery.com/packages/Microsoft.Graph.Groups/)
   - [Microsoft.Graph.Identity.Governance](https://www.powershellgallery.com/packages/Microsoft.Graph.Identity.Governance/)
   - (The runbook has been developed and tested with v2.8.0 of the Microsoft.Graph PowerShell modules.)

4. Decide whether to use a managed identity (recommended) or a custom app registration for authentication to the target Entra ID tenant:
   - Managed identity: Create either a system-assigned or user-assigned managed identity and assign it to the previously configured Azure Automation account according to the [Microsoft documentation](https://learn.microsoft.com/en-us/azure/automation/quickstarts/enable-managed-identity).
   - App registration: Create a custom app registration with a certificate according to the [Microsoft documentation](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app) and upload the app's certificate to the Azure Automation account according to the [Microsoft documentation](https://learn.microsoft.com/en-us/azure/automation/shared-resources/certificates?tabs=azure-powershell#create-a-new-certificate-with-the-azure-portal).

5. Run the [Set-ServicePrincipalPermissions](Set-ServicePrincipalPermissions.ps1) script with a Global Administrator account to assign appropriate Microsoft Graph API permissions ([Group.Read.All](https://learn.microsoft.com/en-us/graph/permissions-reference#groupreadall) & [AccessReview.ReadWrite.All](https://learn.microsoft.com/en-us/graph/permissions-reference#accessreviewreadwriteall)) to either the managed identity or the app registration.

6. In your Azure Automation account, create a new runbook with name `Start-AccessReviews` and PowerShell runtime version `7.2 (preview)` according to the [Microsoft documentation](https://learn.microsoft.com/en-us/azure/automation/manage-runbooks).

7. Activate verbose logging in the Azure runbook according to the [Microsoft documentation](https://learn.microsoft.com/en-us/azure/automation/automation-runbook-output-and-messages#write-output-to-verbose-stream).

8. Upload the source code of the [Start-AccessReviews.ps1](Start-AccessReviews.ps1) to the Azure runbook and adjust all parameters (e.g. $TenantId, $AuthenticationType, ...) to your needs. Use the examples below as a reference. Publish the runbook.


<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- USAGE EXAMPLES -->
## Usage

In the following, you can find some examples how to use the parameters of the runbook:

1. Start access reviews for private M365 groups and MS Teams teams with a "PRJ-" project prefix using managed identity-based authentication:

    | Runbook parameter name | Exemplary value |
    | -----------------------|-------|
    | AuthenticationType | SystemAssignedManagedIdentity |
    | IncludedGroupsMailNicknamePrefixes | PRJ- |
    | IncludedGroupsVisibility | Private |
    | AccessReviewAdditionalMailText | In case of any questions, please contact the Contoso helpdesk. |

2. Start access reviews for private MS Teams teams, except for teams with an "DEPT-" department prefix or an "X-" prefix, using app-based authentication:

    | Runbook parameter name | Exemplary value |
    | -----------------------|-------|
    | TenantId | aa4201e7-433f-44aa-88be-8eeb2d88a908 |
    | AuthenticationType | AppRegistration |
    | AuthenticationAppRegistrationClientId | c091754f-c31c-4d60-ae79-53cfcd3a5d97 |
    | AuthenticationAppRegistrationCertificateThumbprint | 31ED9B5036A479CE2E7180715CC232359D5E8F98 |
    | IncludedGroupsServerSideFilterQuery | groupTypes/any(c:c eq 'Unified') and resourceProvisioningOptions/Any(x:x eq 'Team') |
    | ExcludedGroupsMailNicknamePrefixes | DEPT-;X- |
    | IncludedGroupsVisibility | Private |

3. Only start a "dry run" to check, which groups/teams would be in scope of your access review, with the following command:

    | Runbook parameter name | Exemplary value |
    | -----------------------|-------|
    | AuthenticationType | SystemAssignedManagedIdentity |
    | DryRun | true |

4. Only start a an access review for a single Microsoft 365 group:

    | Runbook parameter name | Exemplary value |
    | -----------------------|-------|
    | AuthenticationType | SystemAssignedManagedIdentity |
    | GroupId | 3c6abeea-a4be-4633-acf6-18927400aeef |


If you want to start the access reviews with certain parameters periodically, e.g. every six months, you can leverage Azure runbook schedules according to the [Microsoft documentation](https://learn.microsoft.com/en-us/azure/automation/shared-resources/schedules).


<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- CONTRIBUTING -->
## Contributing

Contributions are what make the open source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".
Don't forget to give the project a star! Thanks again!

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- CONTACT -->
## Contact

Dustin Schutzeichel - [https://cloudprotect.ninja](https://cloudprotect.ninja)

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->
[contributors-shield]: https://img.shields.io/github/contributors/CloudProtectNinja/MicrosoftCloudSecurity.svg?style=for-the-badge
[contributors-url]: https://github.com/CloudProtectNinja/MicrosoftCloudSecurity/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/CloudProtectNinja/MicrosoftCloudSecurity.svg?style=for-the-badge
[forks-url]: https://github.com/CloudProtectNinja/MicrosoftCloudSecurity/network/members
[stars-shield]: https://img.shields.io/github/stars/CloudProtectNinja/MicrosoftCloudSecurity.svg?style=for-the-badge
[stars-url]: https://github.com/CloudProtectNinja/MicrosoftCloudSecurity/stargazers
[issues-shield]: https://img.shields.io/github/issues/CloudProtectNinja/MicrosoftCloudSecurity.svg?style=for-the-badge
[issues-url]: https://github.com/CloudProtectNinja/MicrosoftCloudSecurity/issues
[license-shield]: https://img.shields.io/github/license/CloudProtectNinja/MicrosoftCloudSecurity.svg?style=for-the-badge
[license-url]: https://github.com/CloudProtectNinja/MicrosoftCloudSecurity/blob/master/LICENSE
[blog-post-url]: https://cloudprotect.ninja/access-reviews-for-specific-m365-groups-and-teams/
