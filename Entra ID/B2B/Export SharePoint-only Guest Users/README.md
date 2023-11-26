<a name="readme-top"></a>



<!-- PROJECT SHIELDS -->
[![Contributors][contributors-shield]][contributors-url]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![MIT License][license-shield]][license-url]



<!-- TITLE -->
# Reporting on SharePoint-Only Guest Users across all SharePoint and OneDrive sites in Microsoft 365



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

The PowerShell script [Export-SharePointOnlyGuestUsers.ps1](Export-SharePointOnlyGuestUsers.ps1) creates a CSV export of all guest users across SharePoint Online and/or OneDrive for Business sites in your Microsoft 365 tenant. The output allows to distinguish between "SharePoint-only guest users" and Entra B2B guest users.

Please find more details in my blog post:

[Reporting on SharePoint-Only Guest Users before Enabling the Entra B2B Integration - Cloud Protect Ninja][blog-post-url]

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- PREREQUISITES -->
## Prerequisites

In order to run the PowerShell script, the following requirements must be met:

1. Install the [PnP PowerShell](https://pnp.github.io/powershell/articles/installation.html) module.
   - The script has been developed and tested with version [1.12.0](https://github.com/pnp/powershell/releases/tag/v1.12.0) of PnP PowerShell.

2. Create an [app registration](https://pnp.github.io/powershell/articles/authentication.html#register-your-own-azure-ad-app) in Entra ID.
   - Configure `Sites.FullControl.All` SharePoint application permissions.

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- USAGE EXAMPLES -->
## Usage

In the following, you can find some examples how to run the PowerShell script.

Create export on guest users across both SharePoint and OneDrive:
```
.\Export-SharePointOnlyGuestUsers.ps1 -TenantName "yourtenant" -ClientId "fa1a81f1-e729-44d8-bb71-0a0c339c0f62" -CertificateThumbprint "789123ABC072E4125785GH4F836AFB12FA64DB210"
```

Create export on guest users across OneDrive only:
```
.\Export-SharePointOnlyGuestUsers.ps1 -IncludedServices @("OneDrive") -TenantName "yourtenant" -ClientId "fa1a81f1-e729-44d8-bb71-0a0c339c0f62" -CertificateThumbprint "789123ABC072E4125785GH4F836AFB12FA64DB210"
```

Create export on guest users across the sites specified in the provided CSV file (with "SiteUrl" column):
```
.\Export-SharePointOnlyGuestUsers.ps1 -SiteUrlInputFile ".\Sites.csv" -TenantName "yourtenant" -ClientId "fa1a81f1-e729-44d8-bb71-0a0c339c0f62" -CertificateThumbprint "789123ABC072E4125785GH4F836AFB12FA64DB210"
```

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
[blog-post-url]: https://cloudprotect.ninja/reporting-sharepoint-only-guest-users/
