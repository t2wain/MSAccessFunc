using MSAccessLib;
using System.Management.Automation;

namespace DAOCmdlets
{
    /// <summary>
    /// Import tables from the source DB to a destination DB.
    /// </summary>
    [Cmdlet(VerbsCommon.Add, "DaoImportTables")]
    public class AddImportTables : Cmdlet, IDisposable
    {
        DB _dbUtil = null!;
        Context _ctx = null!;

        [Parameter(Mandatory = true, Position = 0,
            HelpMessage = "Valid DAO connection string of a destination Access DB.")]
        public string DestConnectString { get; set; } = null!;

        [Parameter(Mandatory = true, Position = 1,
            HelpMessage = "Valid DAO connection string of a source DB")]
        public string SrcConnectString { get; set; } = null!;

        [Parameter(HelpMessage = "Password of source DB")]
        public string SrcPassword { get; set; } = "";

        [Parameter(HelpMessage = "Given a source TableDef and return true for selected TableDef")]
        public ScriptBlock? TableFilter { get; set; }

        [Parameter(HelpMessage = "Given a source TableDef and return a name for destination TableDef")]
        public ScriptBlock? GetDestTableName { get; set; }

        protected override void BeginProcessing()
        {
            _dbUtil = new DB();
            _ctx = this.BuildContext(TableFilter, null, GetDestTableName);
        }

        protected override void ProcessRecord() =>
            _dbUtil.ImportTables(SrcConnectString, DestConnectString, _ctx);

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
