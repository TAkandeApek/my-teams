# my-teams

# Descripition
A webpart that uses Microsoft Graph to list the Teams the current user is a member of. When the user clicks on one of the teams, the web part retrieves information about the default channel (General) and opens it. The web part can be configured to open the team on the web browser or client app.

## Features

- Lists teams the current user is a member of
- Search for teams
- Sort teams in alphabetical order (Ascending/Descending) 
- Sort teams in orderof the created date (Ascending/Descending)
- List channels under a team
- Open a channel directly via link 
- Pagination
- Change item number per page

![Demo](/assets/preview.gif)

## Applies to

- [SharePoint Framework](https:/dev.office.com/sharepoint)

## Prerequisites

- Office 365 subscription with SharePoint Online license

## Version history

| Version | Date              | Comments        |
| ------- | ----------------- | --------------- |
| 1.0     | August 2, 2020 | Initial release |


## Installation Instruction

Clone this repository.

Upload my-teams.sppkg to your SharePoint App Catalog

Once installed, the solution will request the required permissions via the Office 365 admin portal.

If you prefer to approve the permissions in advance, you can do so using Office 365 CLI:

```bash
o365 spo login https://contoso-admin.sharepoint.com
o365 spo serviceprincipal grant add --resource 'Microsoft Graph' --scope 'User.Read.All'
o365 spo serviceprincipal grant add --resource 'Microsoft Graph' --scope 'User.ReadWrite.All'
o365 spo serviceprincipal grant add --resource 'Microsoft Graph' --scope 'Group.Read.All'
o365 spo serviceprincipal grant add --resource 'Microsoft Graph' --scope 'Group.ReadWrite.All'
```

