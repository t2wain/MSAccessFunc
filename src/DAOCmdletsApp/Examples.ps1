# Load the custom module which has 3 cmdlets
Import-Module ".\DAOCmdlets.dll"

# Write out information of the DB.
Out-DaoDbInfo -ConnectString "c:\dev\TestLinkDB.accdb" -InformationAction Continue `
	-TableFilter {param($t) $t.Name -eq "Cable"} | Out-File -FilePath "c:\dev\TestLinkDBInfo.txt"

# Link tables in source DB to destination DB.
Add-DaoImportTables -SrcConnectString "c:\dev\Routing.accdb" -DestConnectString "c:\dev\RoutingImport.accdb" `
	-TableFilter {param($t) $t.Attributes -eq 0} -InformationAction Continue

# Import tables in source DB to destination DB.
Add-DaoLinkTables -SrcConnectString "c:\dev\Routing.accdb" -DestConnectString "c:\dev\RoutingLink.accdb" `
	-InformationAction Continue -TableFilter {param($t) $t.Attributes -eq 0}