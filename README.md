---
page_type: sample
products:
- office-sp
- office-teams
- office-planner
- ms-graph
languages:
- csharp
extensions:
  contentType: samples
  technologies:
  - Microsoft Graph
  - Azure AD
  services:
  - SharePoint
  - Microsoft Teams
  - Planner
  createdDate: 09/12/2018 0:00:00 PM
---
# Contoso Airlines Flight Team Provisioning Sample

This sample app implements Azure functions designed to be invoked via a Graph webhook to provision a Microsoft Team when a new flight is added to a master list in SharePoint. The sample uses Microsoft Graph to do the following provisioning tasks:

- Creates a unified [group](https://docs.microsoft.com/graph/api/resources/groups-overview?view=graph-rest-beta) for the flight team, and initializes a [Team](https://docs.microsoft.com/graph/api/resources/teams-api-overview?view=graph-rest-beta) for the group.
- Creates [channels](https://docs.microsoft.com/graph/api/resources/channel?view=graph-rest-beta) in the team.
- [Installs an app](https://docs.microsoft.com/graph/api/resources/teamsapp?view=graph-rest-beta) to the team.
- Creates a custom SharePoint page and custom [SharePoint list](https://docs.microsoft.com/graph/api/resources/list?view=graph-rest-beta) for the team.
- Adds a [tab](https://docs.microsoft.com/graph/api/resources/teamstab?view=graph-rest-beta) to the team's General channel for the planner plan and SharePoint page.
- [Sends a Graph notification](https://docs.microsoft.com/graph/api/resources/projectrome-notification?view=graph-rest-beta) when the flight is updated.
- [Archives the team](https://docs.microsoft.com/graph/api/team-archive?view=graph-rest-beta) when the flight is deleted.

## Prerequisites

- Visual Studio Code with **Azure Functions** extension installed.
- Office 365 tenant
- Azure subscription if you want to publish the functions. You can run this locally in Visual Studio Code but will need further requirements.

### Prerequisites to run locally

- ngrok
- Azure Cosmos DB Emulator
- Azure Storage Emulator

## Setup

To setup the sample, see [Set up for end-to-end demo](SETUP.md)

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.