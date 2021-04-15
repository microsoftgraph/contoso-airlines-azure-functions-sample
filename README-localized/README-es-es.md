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
# Ejemplo de aprovisionamiento de Contoso Airlines Flight Team

Esta aplicación de ejemplo implementa las funciones de Azure diseñadas para invocarse a través de un webhook de Graph con el fin de aprovisionar un Equipo de Microsoft cuando se agrega un nuevo piloto a una lista maestra en SharePoint. El ejemplo usa Microsoft Graph para realizar las siguientes tareas de aprovisionamiento:

- Crea un [grupo](https://docs.microsoft.com/graph/api/resources/groups-overview?view=graph-rest-beta) unificado para el equipo de vuelo y, a continuación, inicializa un [Equipo](https://docs.microsoft.com/graph/api/resources/teams-api-overview?view=graph-rest-beta) para el grupo.
- Crea [canales](https://docs.microsoft.com/graph/api/resources/channel?view=graph-rest-beta) en el equipo.
- [Instala una aplicación](https://docs.microsoft.com/graph/api/resources/teamsapp?view=graph-rest-beta) en un equipo.
- Crea una página de SharePoint personalizada y una](https://docs.microsoft.com/graph/api/resources/list?view=graph-rest-beta)lista de SharePoint[ personalizada para el equipo.
- Agrega una [pestaña](https://docs.microsoft.com/graph/api/resources/teamstab?view=graph-rest-beta) al canal General del Equipo para el plan de Planner y la página de SharePoint.
- [Envía una notificación de Graph](https://docs.microsoft.com/graph/api/resources/projectrome-notification?view=graph-rest-beta) cuando se actualiza el vuelo.
- [Archiva el equipo](https://docs.microsoft.com/graph/api/team-archive?view=graph-rest-beta) cuando se elimina el vuelo.

## Requisitos previos

- Extensión de Visual Studio Code con **funciones de Azure** instalada.
- Inquilino de Office 365
- Suscripción de Azure si desea publicar las funciones. Puede ejecutarlo de forma local en Visual Studio Code, pero necesitará requisitos adicionales.

### Requisitos previos para ejecutar localmente

- ngrok
- Azure Cosmos DB Emulator
- Azure Storage Emulator

## Instalación

Para configurar el ejemplo, consulte [Configurar la demostración de un extremo a otro](SETUP.md)

Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/).
Para obtener más información, vea [Preguntas frecuentes sobre el código de conducta](https://opensource.microsoft.com/codeofconduct/faq/)
o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.