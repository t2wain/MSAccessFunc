using MSAccessLib;
using System.Management.Automation;

namespace DAOCmdlets
{
    /// <summary>
    /// Link tables from a source DB to a destination DB.
    /// </summary>
    [Cmdlet(VerbsCommon.Add, "DaoLinkTables")]
    public class AddLinkTables : Cmdlet, IDisposable
    {
        DB _dbUtil = null!;
        Context _ctx = null!;

        [Parameter(Mandatory = true,
            HelpMessage = "Valid DAO connection string of a destination Access DB.")]
        public string DestConnectString { get; set; } = null!;

        [Parameter(Mandatory = true,
            HelpMessage = "Valid DAO connection string of a source DB")]
        public string SrcConnectString { get; set; } = null!;

        [Parameter(HelpMessage = "Password of source DB")]
        public string SrcPassword { get; set; } = "";

        [Parameter(HelpMessage = "Save source DB password with linked tables in destination DB")]
        public SwitchParameter SavePassword { get; set; }

        [Parameter(HelpMessage = "Given a source TableDef and return true for selected TableDef ( {param($t) $true} )")]
        public ScriptBlock? TableFilter { get; set; }

        [Parameter(HelpMessage = "Given a source TableDef and return a new name for destination TableDef ( {param($t) $t.Name} )")]
        public ScriptBlock? GetDestTableName { get; set; }

        protected override void BeginProcessing()
        {
            _dbUtil = new DB();
            _ctx = this.BuildContext(TableFilter, null, GetDestTableName) with
            {
                IsSavePwdWithLinkTable = SavePassword
            };
        }

        protected override void ProcessRecord() =>
            _dbUtil.LinkTables(SrcConnectString, DestConnectString, _ctx);

        protected override void EndProcessing() =>
            Dispose();

        protected override void StopProcessing() =>
            Dispose();


        public void Dispose()
        {
            _ctx?.Writer?.Flush();
            _ctx?.Dispose();
            _dbUtil?.Dispose();
            _ctx = null!;
            _dbUtil = null!;
        }

    }
}
