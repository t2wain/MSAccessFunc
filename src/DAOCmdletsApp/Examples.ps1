

Out-DaoDbInfo -ConnectString "c:\dev\TestLinkDB.accdb" -InformationAction Continue `
	-TableFilter {param($t) $t.Name -eq "Cable"} | Out-File -FilePath "c:\dev\TestLinkDBInfo.txt"