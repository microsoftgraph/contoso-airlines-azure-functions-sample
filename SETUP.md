# Deploy solution in an Office 365 tenant

WORK IN PROGRESS

## Demo users

- **Flight administrator**: An admin user that you'll use to demo creating flights in the SharePoint list.
- **Flight attendant**: A non-admin user that you'll use to demo the mobile app and the self-serve aspect of the web app.

## Create groups

Create the following universal groups and add folks to them.

- **Flight Admins**: Add the **Flight administrator** account.
- **Flight Attendants**: Add the **Flight attendant** account + 4-5 others

## Create flight team admin site

1. Create a team site in SharePoint named `FlightAdmin`. (**Note**: If you use a different site name, be sure to update the `FlightAdminSite` value in **local.settings.json**.)
1. Create a new document library in the site named `Flights`. (**Note**: If you use a different document library name, be sure to update the `FlightList` value in **local.settings.json**.)
1. Add the following columns to the document library.
    1. **Single line of text**

        | Field                                         | Value        |
        |-----------------------------------------------|--------------|
        | Name                                          | Description  |
        | Require that this column contains information | Yes          |

    1. **Number**

        | Field                                         | Value         |
        |-----------------------------------------------|---------------|
        | Name                                          | Flight Number |
        | Number of decimal places                      | 0             |
        | Require that this column contains information | Yes           |
        | Enforce unique values                         | Yes           |

    1. **Person**

        | Field                                         | Value  |
        |-----------------------------------------------|--------|
        | Name                                          | Pilots |
        | Allow multiple selections                     | Yes    |
        | Require that this column contains information | Yes    |

    1. **Person**

        | Field                                         | Value             |
        |-----------------------------------------------|-------------------|
        | Name                                          | Flight Attendants |
        | Allow multiple selections                     | Yes               |
        | Require that this column contains information | No                |

    1. **Single line of text**

        | Field                                         | Value            |
        |-----------------------------------------------|------------------|
        | Name                                          | Catering Liaison |
        | Require that this column contains information | No               |

    1. **Date**

        | Field                                         | Value          |
        |-----------------------------------------------|----------------|
        | Name                                          | Departure Time |
        | Include Time                                  | Yes            |
        | Require that this column contains information | Yes            |

    1. **Single line of text**

        | Field                                         | Value          |
        |-----------------------------------------------|----------------|
        | Name                                          | Departure Gate |
        | Require that this column contains information | Yes            |

1. Select the **New** dropdown, then select **Edit New menu**. Disable all items except **Word document**, then select **Save**.
1. Select the gear icon in the upper right, then select **Library settings**.
1. Select **Indexed columns**, then select **Create a new index**.
1. Set **Primary column for this index** to **Departure Time** then select **Create**.

## App registration

Register an app **Flight Team Provisioning Function**.

- Accounts in this organizational directory only
- Redirect URI: Web, https://flights.contoso.com
- Add application permissions for Graph:
  - **Calendars.ReadWrite**
  - **Files.ReadWrite.All**
  - **Group.ReadWrite.All**
  - **Sites.Manage.All**
  - **Sites.ReadWrite.All**
  - **User.Invite.All**
  - **User.Read.All**
- After adding the permissions, use the **Grant admin consent for Contoso** button
- Create a secret
- Set `TenantId`, `TenantName`, `AppId`, and `AppSecret`

## OPTIONAL: Configuring Graph notifications

This section deals with the configuration needed to enable the [Graph notifications](https://docs.microsoft.com/graph/api/resources/notifications-api-overview?view=graph-rest-beta) feature of this sample. If you do not do these steps, the sample will still work, it just will not send these notifications.

You'll need a few things for this to work:

- An application to receive the notifications. This sample was written to work with the [Contoso Airlines iOS app](https://github.com/microsoftgraph/contoso-airlines-ios-swift-sample), but you could also write your own. For the sample iOS app, you need:
  - A MacOS device with XCode installed.
  - An Apple developer account
- You must [register your app in the Windows Dev Center for cross-device experiences](https://docs.microsoft.com/windows/project-rome/notifications/how-to-guide-for-ios#register-your-app-in-microsoft-windows-dev-center-for-cross-device-experiences).

### App registration for notification service

Register a separate app in the Azure portal for the notification service named **Flight Team Notification Service**.

- Accounts in this organizational directory only
- Redirect URI: Web, https://flights.contoso.com
- Add delegated permissions for Graph:
  - **Notifications.ReadWrite.CreatedByApp**
  - **User.Read**
  - **offline_access**
- After adding the permissions, use the **Grant admin consent for Contoso** button
- On **Expose an API** tab, add a scope named **Notifications.Send** that admins and users can consent. Accept the application ID URI that is generated for you.
- Add an **Authorized client application** using the application ID for your receiving app (for example, the iOS sample above)
- Create a secret
- Set the values for `NotificationAppId` and `NotificationAppSecret` in **local.settings.json**.
- Add the application ID for **Flight Team Notification Service** in the **Support Microsoft Account & Azure Active Directory** section of your cross-device experience registration in the Windows Dev Center.

### Add your cross-device app domain

Set the `NotificationHostName` value in **local.settings.json** to the domain configured in the **Verify your cross-device app domain** section of your cross-device experience registration in the Windows Dev Center.