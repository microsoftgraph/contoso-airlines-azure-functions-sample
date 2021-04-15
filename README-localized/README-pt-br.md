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
# Exemplo de provisionamento da equipe de bordo da Contoso Airlines

Este exemplo de aplicativo implementa funções do Azure desenvolvidas para serem chamadas por meio de um webhook do Graph para provisionar um Microsoft Team quando um novo voo é adicionado a uma lista mestre no SharePoint. O exemplo usa o Microsoft Graph para executar as seguintes tarefas de provisionamento:

- Criar um [grupo](https://docs.microsoft.com/graph/api/resources/groups-overview?view=graph-rest-beta) unificado para a equipe de voo e inicializar um [Team](https://docs.microsoft.com/graph/api/resources/teams-api-overview?view=graph-rest-beta) para o grupo.
- Criar [canais](https://docs.microsoft.com/graph/api/resources/channel?view=graph-rest-beta) na equipe.
- [Instalar um aplicativo](https://docs.microsoft.com/graph/api/resources/teamsapp?view=graph-rest-beta) para a equipe.
- Criar uma página do SharePoint personalizada e uma [lista do SharePoint](https://docs.microsoft.com/graph/api/resources/list?view=graph-rest-beta) personalizada para a equipe.
- Adicionar uma [guia](https://docs.microsoft.com/graph/api/resources/teamstab?view=graph-rest-beta) ao canal "Geral" da equipe para o plano do Planner e a página do SharePoint.
- [Enviar uma notificação do Graph](https://docs.microsoft.com/graph/api/resources/projectrome-notification?view=graph-rest-beta) quando o voo for atualizado.
- [Arquivar a equipe](https://docs.microsoft.com/graph/api/team-archive?view=graph-rest-beta) quando o voo for excluído.

## Pré-requisitos

- Código do Visual Studio com a extensão **Azure Functions** instalada.
- Locatário do Office 365
- Assinatura do Azure para publicar as funções. Você pode executar isso localmente no Código do Visual Studio, mas precisará atender a mais requisitos.

### Pré-requisitos para executar localmente

- ngrok
- Emulador de banco de dados do Azure Cosmos
- Emulador de armazenamento do Azure

## Configuração

Para configurar ao exemplo, confira [Configurar para demonstração de ponta a ponta](SETUP.md)

Este projeto adotou o [Código de Conduta de Código Aberto da Microsoft](https://opensource.microsoft.com/codeofconduct/).
Para saber mais, confira as [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/)
ou entre em contato pelo [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.