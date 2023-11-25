# react-hello-m-365

## Summary

This is the solution sample used for my session at CollabDays Lisbon 2023

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.18.1-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Extend Teams across Microsoft 365](https://learn.microsoft.com/en-us/microsoftteams/platform/m365-apps/overview)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

### Run tests with Playwright

In order to run your tests locally & remotely, both with the VS Code extension and the CLI (local & through GitHub Actions), you need to configure environment variables.

#### For the VS Code extension

You need to add the following settings in the `.vscode/settings.json` file:

```json
"playwright.env": {
    "M365_LOGIN": "",
    "M365_PWD": ""
  }
```

#### For the CLI

You need to add a `.env` file at the root of your project, which will contain following parameters:

```
M365_LOGIN=''
M365_PWD=''
PLAYWRIGHT_SERVICE_ACCESS_TOKEN=''
PLAYWRIGHT_SERVICE_URL=''
```

#### GitHub Actions

Go to **Settings**, then **Secrets** and add the following *repository secrets*:

* M365_LOGIN
* M365_PWD

If you plan to use Microsoft Playwright, add:

* PLAYWRIGHT_SERVICE_ACCESS_TOKEN
* PLAYWRIGHT_SERVICE_URL
* PLAYWRIGHT_SERVICE_RUN_ID

If you plan to publish Playwright reports on Azure, add:

* AZCOPY_SPA_APPLICATION_ID
* AZCOPY_SPA_CLIENT_SECRET
* AZCOPY_TENANT_ID

### Context testing authentication

To ensure that your tests will work in a real context, you need to be authenticated. So you need to provide an account in the environment variables (and the GitHub Action secrets). But this account must be out of the common user security rules, such as MFA. You can configure this one (you or with your Ops) to be secured anyway. For example, you can scope its location to the same region as the pipeline will run in order not to trigger geographic alert.

### Teams personal apps

Authenticate to [Teams Admin Portal](https://admin.teams.microsoft.com) in order to enable personal apps.

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| react-hello-m365 | Michaël Maillot, onepoint, [twitter](https://twitter.com/michael_maillot) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | November 25, 2023 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp bundle**
  - **gulp package-solution**
- Grab the SPFx package in the `sharepoint/solution` folder and drop it in the Tenant App Catalog
- Enable the deployment on the whole tenant and sync it to Teams
- Authenticate to Microsoft 365 with the account configured in the environment variable
- Add the app called **Hello M365** in Teams, Outlook and Microsoft 365 portal
- In the `tests\m365app.spec.ts`, configure the following constant:
  - **APP_ID**: the Teams personnal app ID (you can find it on the Teams manifest file if you've generated one once deployed on Teams if auto-generated)

Once environment variables created and configured, you can run the following commands:

`npx playwright test`


For testing Microsoft Playwright Testing: 

`npx playwright test m365app --config=playwright.service.config.ts --workers=20`

## Features

- Basic tests including authentication to following platforms:
  - Teams
  - Outlook
  - Microsoft 365 portal
- TeamsJS SDK capabilities tests included
- GitHub Action ready to use with
  - Microsoft Playwright (Azure subscription required)
  - Azure Static Website for Playwright reporting 

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [Playwright documentation](https://playwright.dev/docs/api/class-playwright)
- [Microsoft Playwright Testing](https://azure.microsoft.com/en-us/products/playwright-testing/)
- [Elio Struyf articles about Playwright](https://www.eliostruyf.com/tags/playwright/)
