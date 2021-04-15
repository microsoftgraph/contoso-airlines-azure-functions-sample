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
# Contoso 航空フライト チームのプロビジョニング サンプル

このサンプル アプリは、SharePoint のマスター リストに新しいフライトが追加されると、Microsoft Teams チームをプロビジョニングするために Graph webhook 経由で呼び出されるように作られた Azure 関数を実装します。このサンプルでは、Microsoft Graph を使用して次のプロビジョニング タスクを実行します。

- 統合された[グループ](https://docs.microsoft.com/graph/api/resources/groups-overview?view=graph-rest-beta)をフライト チーム用に作成し、そのグループについて [Teams](https://docs.microsoft.com/graph/api/resources/teams-api-overview?view=graph-rest-beta) を初期化します。
- チームに[チャネル](https://docs.microsoft.com/graph/api/resources/channel?view=graph-rest-beta) を作成します。
- チームに[アプリをインストール](https://docs.microsoft.com/graph/api/resources/teamsapp?view=graph-rest-beta)します。
- チームにカスタム SharePoint ページとカスタム [SharePoint](https://docs.microsoft.com/graph/api/resources/list?view=graph-rest-beta) リストを作成します。
- プランナーの計画と SharePoint ページ用の[タブ](https://docs.microsoft.com/graph/api/resources/teamstab?view=graph-rest-beta)をチームの \[一般] チャネルに追加します。
- フライトが更新されたときに、[Graph の通知を送信](https://docs.microsoft.com/graph/api/resources/projectrome-notification?view=graph-rest-beta)します。
- フライトが削除されたときに、[チームをアーカイブ](https://docs.microsoft.com/graph/api/team-archive?view=graph-rest-beta)します。

## 前提条件

- **Azure Functions** 拡張機能がインストールされている Visual Studio Code。
- Office 365 テナント
- 関数を発行する場合は、Azure のサブスクリプション。このアプリは Visual Studio Code でローカル実行できますが、その場合は追加の要件があります。

### ローカルで実行するための前提条件

- ngrok
- Azure Cosmos DB Emulator
- Azure Storage Emulator

## セットアップ

サンプルをセットアップするには、「[Set up for end-to-end demo (エンド ツー エンド デモ用のセットアップ)](SETUP.md)」を参照してください。

このプロジェクトでは、[Microsoft Open Source Code of Conduct (Microsoft オープン ソース倫理規定)](https://opensource.microsoft.com/codeofconduct/)
が採用されています。詳細については、「[Code of Conduct の FAQ (倫理規定の FAQ)](https://opensource.microsoft.com/codeofconduct/faq/)」
を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。