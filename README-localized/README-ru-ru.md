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
# Пример подготовки летного состава для авиакомпании Contoso

В этом примере приложения реализованы функции Azure, рассчитанные на вызов через веб-перехватчик Graph для подготовки команды Microsoft Teams при добавлении нового рейса в главный список в SharePoint. В этом примере приложение Microsoft Graph используется для выполнения задач подготовки, перечисленных ниже.

- Создание единой[группы](https://docs.microsoft.com/graph/api/resources/groups-overview?view=graph-rest-beta) для летного экипажа и инициализация [команды](https://docs.microsoft.com/graph/api/resources/teams-api-overview?view=graph-rest-beta) для этой группы.
- Создание [каналов](https://docs.microsoft.com/graph/api/resources/channel?view=graph-rest-beta) в команде.
- [Установка приложения](https://docs.microsoft.com/graph/api/resources/teamsapp?view=graph-rest-beta) для команды.
- Создание настраиваемой страницы SharePoint и настраиваемого [списка SharePoint](https://docs.microsoft.com/graph/api/resources/list?view=graph-rest-beta) для команды.
- Добавление [вкладки](https://docs.microsoft.com/graph/api/resources/teamstab?view=graph-rest-beta) для плана Планировщика и страницы SharePoint в "Общий" канал.
- [Отправка уведомления Graph](https://docs.microsoft.com/graph/api/resources/projectrome-notification?view=graph-rest-beta) при обновлении информации о рейсе.
- [Архивирование команды](https://docs.microsoft.com/graph/api/team-archive?view=graph-rest-beta) при удалении информации о рейсе.

## Необходимые компоненты

- Visual Studio Code с установленным расширением **Функции Azure**.
- Клиент Office 365
- Подписка на Azure, если нужна публикация функций. Это можно сделать локально в приложении Visual Studio Code, но для этого потребуется выполнить дополнительные требования.

### Предварительные условия для локального выполнения

- ngrok
- Эмулятор Azure Cosmos DB
- Эмулятор службы хранилища Azure

## Настройка

Информацию о настройке примера приложения см. в статье [Настройка полной демонстрации](SETUP.md)

Этот проект соответствует [Правилам поведения разработчиков открытого кода Майкрософт](https://opensource.microsoft.com/codeofconduct/).
Дополнительные сведения см. в разделе [часто задаваемых вопросов о правилах поведения](https://opensource.microsoft.com/codeofconduct/faq/).
Если у вас возникли вопросы или замечания, напишите нам по адресу [opencode@microsoft.com](mailto:opencode@microsoft.com).