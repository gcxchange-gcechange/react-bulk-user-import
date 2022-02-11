# react-bulk-user-import

## Summary

Simple admin form to queue a Sharepoint list up for bulk import.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.11-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> This webpart works with a [custom function app](https://github.com/gcxchange-gcechange/appsvc_fnc_bulkuserimport) and a specific Sharepoint list format.

## Solution

Solution|Author(s)
--------|---------
bulk-user-import | piet0024

## Version history

Version|Date|Comments
-------|----|--------
1.0|January 19, 2022|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Be sure to add your client ID and URL to your [function app](https://github.com/gcxchange-gcechange/appsvc_fnc_bulkuserimport) in the `sendImportQueue` function

> Be sure to update the `webApiPermissionRequests` in `config/package-solution` to align with your function app

## Features

A simple form that an admin can put in the name of a sharepoint list and queue the bulk user import.

The webpart will give feedback if the queue is starting or if the list was not found.
## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
