using MSAccessLib;
using System.Management.Automation;

namespace DAOCmdlets
{
    /// <summary>
    /// Write out information about the database.
    /// </summary>
    [Cmdlet(VerbsData.Out, "DaoDbInfo")]
    public class OutDbInfo : Cmdlet, IDisposable
    {
        DB _dbUtil = null!;
        Context _ctx = null!;

        [Parameter(Mandatory = true, Position = 0,
            HelpMessage = "Valid DAO connection string of a DB")]
        public string ConnectString { get; set; } = null!;

        [Parameter(HelpMessage = "Given a TableDef and return true for selected TableDef")]
        public ScriptBlock? TableFilter { get; set; }

        [Parameter(HelpMessage = "Given a QueryDef and return true for selected QueryDef")]
        public ScriptBlock? QueryFilter { get; set; }

        [Parameter(HelpMessage = "Do not write out empty properties.")]
        public SwitchParameter HideEmptyProperty { get; set; }

        [Parameter(HelpMessage = "Do not write out properties of Field.")]
        public SwitchParameter HideFieldProperty { get; set; }

        protected override void BeginProcessing()
        {
            _dbUtil = new DB();
            _ctx = this.BuildContext(TableFilter, QueryFilter) with
            {
                HideEmptyProperty = HideEmptyProperty,
                HideFieldProperty = HideFieldProperty,
                IsMSAccessDB =
                    ConnectString.Contains(".accdb")
                    || ConnectString.Contains(".mdb")
            };
        }

        protected override void ProcessRecord() =>
            _dbUtil.PrintDatabase(ConnectString, _ctx);

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
