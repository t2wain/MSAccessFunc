## PowerShell Custom Module with DAO

This project implements a custom PowerShell module "DAOCmdlets.dll" with three (3) cmdlets:

- Out-DaoDbInfo
- Add-DaoLinkTables
- Add-DaoImportTables

Please use the help command to show all the parameters.

Note, since this project is a class library project type, not all the dependencies will be output when the project is compiled. There is a do-nothing Console app _DAOCmdletsApp_ that will do just that, collecting all the dependencies when compiled. The custom PowerShell module should be debugged or loaded from the _DAOCmdletsApp_ project.

## Examples

```powershell

# Load the custom module which has 3 cmdlets
Import-Module ".\DAOCmdlets.dll"

# Write out information (tables/columns/queries/index) of the DB.
Out-DaoDbInfo -ConnectString "c:\dev\TestLinkDB.accdb" -InformationAction Continue `
	-TableFilter {param($t) $t.Name -eq "Cable"} | Out-File -FilePath "c:\dev\TestLinkDBInfo.txt"

# Link tables from source DB into destination DB.
Add-DaoImportTables -SrcConnectString "c:\dev\Routing.accdb" -DestConnectString "c:\dev\RoutingImport.accdb" `
	-TableFilter {param($t) $t.Attributes -eq 0} -InformationAction Continue

# Import tables from source DB into destination DB.
Add-DaoLinkTables -SrcConnectString "c:\dev\Routing.accdb" -DestConnectString "c:\dev\RoutingLink.accdb" `
	-InformationAction Continue -TableFilter {param($t) $t.Attributes -eq 0}

# Write out information of the SPELREF Oracle Schema.
Out-DaoDbInfo -ConnectString "ODBC;FILEDSN=C:\devgit\Data\XXX\XXXSPELREF.dsn" `
    -TableFilter {param($t) $t.Name -match "^XXX_XXXELREF\."} `
    -HideFieldProperty -HideEmptyProperty -InformationAction Continue 

# Link tables in SPELREF Oracle Schema to destination MSAccess DB.
Add-DaoLinkTables -SrcConnectString "ODBC;FILEDSN=C:\devgit\Data\XXX\XXXSPELREF.dsn" `
    -DestConnectString "c:\devgit\Data\XXX\XXX_SPELREF.accdb" `
    -TableFilter {param($t) $t.Name -match "^XXX_XXXELREF\."} `
    -GetDestTableName {param($t) $t.Name.Replace("XXX_XXXELREF.", "")} `
    -SavePassword -InformationAction Continue

```

## ILogger and TextWriter vs. PowerShell WriteXXX Methods

The original .NET library _MSAccessLib.dll_ that implements all the DAO methods uses _ILogger_ and _TextWriter_ objects internally. Howerver, in the PowerShell environment, cmdlet uses its own various WriteXXX methods, such as WriteObject and WriteInformation. In this project, the class _CmdletLogger_ and _CmdletTextWriter_ wrap the cmdlet instance and redirect the ILogger and TextWriter method calls to the corresponding cmdlet WriteXXX method calls.

## .NET Pred<> and Func<> vs. PowerShell ScriptBlock

Another method call that requires redirection is the .NET _Pred<>_ and _Func<>_ vs. the PowerShell _ScriptBlock_ property types. Within .NET, the Pred<> and Func<> method calls (and their parameters) be forwarded to the ScriptBlock calls and the return value be sent back to .NET. 