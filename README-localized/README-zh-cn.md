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
# Contoso Airlines 航班团队预配示例

此示例应用实现旨在通过 Graph Webhook 调用的 Azure 函数，用于在有新航班添加到 SharePoint 母版列表时预配 Microsoft 团队。此示例使用 Microsoft Graph 执行以下预配任务：

- 为航班团队创建统一的[组](https://docs.microsoft.com/graph/api/resources/groups-overview?view=graph-rest-beta)，并为该组初始化一个[团队](https://docs.microsoft.com/graph/api/resources/teams-api-overview?view=graph-rest-beta)。
- 在团队中创建[频道](https://docs.microsoft.com/graph/api/resources/channel?view=graph-rest-beta)。
- 向团队[安装应用程序](https://docs.microsoft.com/graph/api/resources/teamsapp?view=graph-rest-beta)。
- 为团队创建自定义 SharePoint 页面和自定义 [SharePoint 列表](https://docs.microsoft.com/graph/api/resources/list?view=graph-rest-beta)。
- 向规划器计划和 SharePoint 页面的团队“常规”频道添加[选项卡](https://docs.microsoft.com/graph/api/resources/teamstab?view=graph-rest-beta)。
- 在更新航班时[发送 Graph 通知](https://docs.microsoft.com/graph/api/resources/projectrome-notification?view=graph-rest-beta)。
- 在删除航班时[将团队存档](https://docs.microsoft.com/graph/api/team-archive?view=graph-rest-beta)。

## 先决条件

- 装有 **Azure Functions** 扩展的 Visual Studio Code。
- Office 365 租户
- Azure 订阅 - 如果希望发布函数。可在 Visual Studio 代码中本地运行此操作，但需要满足更多要求。

### 本地运行的先决条件

- ngrok
- Azure Cosmos DB 模拟器
- Azure 存储模拟器

## 设置

若要设置示例，请参阅[进行端到端演示设置](SETUP.md)

此项目遵循 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。
有关详细信息，请参阅[行为准则常见问题解答](https://opensource.microsoft.com/codeofconduct/faq/)。
如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。