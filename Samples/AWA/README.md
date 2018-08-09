# Access Web App Samples

* Northwind - The result of importing the Access desktop Northwind template into an Access Web App. 

## Using a sample

2. Download the sample .dacpac that you are interested in.

1. Create a new database in SQL Server or SQL Azure.  The Azure portal or SQL Server Management Studio are good tools for this.  Remember the database name you used.

2. Install sqlpackage.exe.  If you have install SQL Server Management Studio, it is already included in that distribution.  Or you can install using these directions: [https://docs.microsoft.com/en-us/sql/tools/sqlpackage-download?view=sql-server-2017](https://docs.microsoft.com/en-us/sql/tools/sqlpackage-download?view=sql-server-2017)

3. Full documentation on sqlpackage is available at [https://docs.microsoft.com/en-us/sql/tools/sqlpackage?view=sql-server-2017](https://docs.microsoft.com/en-us/sql/tools/sqlpackage?view=sql-server-2017)

4. Run

	sqlpackage /a:publish /sf:sample.dacpac /tsn:servername /tdn:databasename /tu:username /tp:userpassword

5. Wait.  It takes a while for the tool to perform its steps.  While it is operating, it will produce a list of actions it is performing and a final "success" message at the end.

6. You now have the schema and data of your AWA moved to SQL.

## Creating a sample

To create your own sample .dacpac from an Access Web App:

1. Open the AWA from Access Desktop.

2. Go to the File Menu, select Save As.

4. Under the default Save Database As, and Save as Snapshot.

5. Name the snapshot and select a file location.  

6. The resulting file will have a .app extension.  Using the Windows explorer, rename this file with a .zip extension.

7. Open the zip file using Windows explorer or another tool.  You will discover within an appdb.dacpac file.  This is the file you seek that contains the AWA's schema and data.   
  

