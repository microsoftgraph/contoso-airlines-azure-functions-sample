# Deploy solution in an Office 365 tenant

WORK IN PROGRESS

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

        | Field        | Value          |
        |--------------|----------------|
        | Name         | Departure Time |
        | Include Time | Yes            |

    1. **Single line of text**

        | Field                                         | Value          |
        |-----------------------------------------------|----------------|
        | Name                                          | Departure Gate |
        | Require that this column contains information | Yes            |

1. Select the **New** dropdown, then select **Edit New menu**. Disable all items except **Word document**, then select **Save**.
