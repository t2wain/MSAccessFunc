# Using Microsoft DAO to build utility methods for Access database

This library provide many utiliy methods for Microsoft Access database using the DAO library

- Connect to remote databases with either a configured DSN, a DSN file, or a connection string (DSN-less)
- Link or import all tables from remote database into Access
- Display all the infos about tables, queries, fields, and indexes

## Why use DAO

I am using DAO just for certain native functionalities specific to Access such as creating link tables and importing entire table data into Access. I am NOT using DAO for general data access. For general data access, I prefer using ADO.NET or EntityFramework.