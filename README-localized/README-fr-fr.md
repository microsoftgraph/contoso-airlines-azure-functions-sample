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
# Exemple d’approvisionnement pour le personnel de bord Contoso Airlines

Cet exemple d’application implémente les fonctions Azure conçues pour être appelées par un webhook pour approvisionner une équipe Microsoft lorsqu’un nouveau vol est ajouté à une liste principale dans SharePoint. L’exemple utilise Microsoft Graph pour effectuer les tâches d'approvisionnement suivantes :

- Crée un [groupe](https://docs.microsoft.com/graph/api/resources/groups-overview?view=graph-rest-beta) unifié pour l’équipe de vol et Initialise une [Équipe](https://docs.microsoft.com/graph/api/resources/teams-api-overview?view=graph-rest-beta) pour le groupe.
- Crée des [canaux](https://docs.microsoft.com/graph/api/resources/channel?view=graph-rest-beta) dans l’équipe.
- [Installe une application](https://docs.microsoft.com/graph/api/resources/teamsapp?view=graph-rest-beta) dans l'équipe.
- Crée une page SharePoint sur mesure et personnalisée la [Liste SharePoint](https://docs.microsoft.com/graph/api/resources/list?view=graph-rest-beta) pour l’équipe.
- Ajoute un [onglet](https://docs.microsoft.com/graph/api/resources/teamstab?view=graph-rest-beta) au canal général de l’équipe pour le plan du planificateur et la page SharePoint.
- [Envoie une notification Graph](https://docs.microsoft.com/graph/api/resources/projectrome-notification?view=graph-rest-beta) le vol est mis à jour.
- [Archive l’équipe](https://docs.microsoft.com/graph/api/team-archive?view=graph-rest-beta) lorsque le vol supprimé.

## Conditions préalables

- Extension Visual Studio Code installée avec **Azure Functions**.
- Client Office 365
- Abonnement Azure si vous souhaitez publier les fonctions. Vous pouvez l’exécuter localement dans Visual Studio Code, mais d'autres conditions sont requises.

### Conditions préalables à l’exécution locale

- ngrok
- Émulateur DB Azure Cosmos DB
- Émulateur de stockage Azure

## Configuration

Pour configurer l’exemple, voir [Configurer pour une démonstration de bout en bout](SETUP.md)

Ce projet a adopté le [Code de conduite Microsoft Open Source](https://opensource.microsoft.com/codeofconduct/).
Pour en savoir plus, consultez la [FAQ relative au Code de conduite](https://opensource.microsoft.com/codeofconduct/faq/)
ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.