# CDSImport
Import the schema and data from an Access Web App into a Common Data Service database.  Understands specifics of the source database and supports relationships and images.

### Overview
**CDSImport** contains one primary commandlet **Import-CDSAccessWebApp**. 

This module builds on top of [Microsoft.Xrm.Data.PowerShell](https://github.com/seanmcne/Microsoft.Xrm.Data.PowerShell), which in turns builds on [Microsoft.Xrm.Tooling.CrmConnector.Powershell](https://docs.microsoft.com/en-us/powershell/module/microsoft.xrm.tooling.crmconnector.powershell/?view=dynamics365ce-ps).  It also uses the [SqlServer](https://docs.microsoft.com/en-us/sql/powershell/sql-server-powershell?view=sql-server-2017) module.  

