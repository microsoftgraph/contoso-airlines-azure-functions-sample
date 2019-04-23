# Contoso Airlines Flight Team Provisioning Sample

This sample app implements Azure functions designed to be invoked via a Graph webhook to provision a Microsoft Team when a new flight is added to a master list in SharePoint. The sample uses Microsoft Graph to do the following provisioning tasks:

- Creates a unified [group](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/groups-overview) for the flight team, and initializes a [Team](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/teams_api_overview) for the group.
- Creates [channels](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/channel) in the team.
- [Installs an app](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/teamsapp) to the team.
- Creates a custom SharePoint page and custom [SharePoint list](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/list) for the team.
- Adds [tabs](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/teamstab) to the team's General channel for the planner plan and SharePoint page.
- [Sends a cross-device notification](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/projectrome_notification) when the flight is updated.
- [Archives the team](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/team_archive) when the flight is deleted.

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