## Microsoft Data Access Object (DAO) Utility for Microsoft Access database

This library provide many methods using Microsoft DAO library

- Connect to remote databases with either a configured DSN, a DSN file, or a connection string (DSN-less)
- Link or import all tables from remote database into Access
- Display all the infos about tables, queries, fields, and indexes

Please view the examples and unit tests on ways to use the library.

## Custom PowerShell module for DAO

This project also implements a custom PowerShell module with 3 cmdlets.

## Why use Microsoft Access?

I often have a need to analyze data of various remote Oracle databases. For complex task, I prefer to use Microsoft Access application to perform the analysis by either linking or importing tables from remote database into Access. Access has the following advantages:

- Access is a full features database application with nice GUI to work with data.
- I can create temporary tables and save queries inside Access without altering the production database.
- I can link/import tables from multiple databases into a single Access file.
- The queries run faster with local imported data rather than across slow network to the remote database.


