using MSAccessLib;
using System.Management.Automation;

namespace DAOCmdlets
{
    [Cmdlet(VerbsCommon.Add, "DaoLinkTables")]
    public class AddLinkTables : Cmdlet
    {
        DB _dbUtil = null!;
        Context _ctx = null!;

        [Parameter(Mandatory = true)]
        public string SrcConnectString { get; set; } = null!;

        [Parameter(Mandatory = true)]
        public string DestConnectString { get; set; } = null!;

        [Parameter]
        public string SrcPassword { get; set; } = "";

        [Parameter]
        public SwitchParameter SavePassword { get; set; }

        [Parameter]
        public ScriptBlock? TableFilter { get; set; }

        [Parameter]
        public ScriptBlock? GetNewTableName { get; set; }

        protected override void BeginProcessing()
        {
            _dbUtil = new DB();
            _ctx = this.BuildContext(TableFilter, null, GetNewTableName) with
            {
                IsSavePwdWithLinkTable = SavePassword
            };
        }

        protected override void ProcessRecord()
        {
            _dbUtil.LinkTables(SrcConnectString, DestConnectString, _ctx);
        }

        protected override void EndProcessing()
        {
            _ctx.Writer.Flush();
            _ctx.Dispose();
            _dbUtil.Dispose();
        }

    }
}
