using MSAccessLib;
using System.Management.Automation;

namespace DAOCmdlets
{
    [Cmdlet(VerbsData.Out, "DaoDbInfo")]
    public class OutDbInfo : Cmdlet
    {
        DB _dbUtil = null!;
        Context _ctx = null!;

        [Parameter(Mandatory = true)]
        public string ConnectString { get; set; } = null!;

        [Parameter]
        public ScriptBlock? TableFilter { get; set; }

        [Parameter]
        public ScriptBlock? QueryFilter { get; set; }

        [Parameter]
        public SwitchParameter HideEmptyProperty { get; set; }

        [Parameter]
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

        protected override void ProcessRecord()
        {
            _dbUtil.PrintDatabase(ConnectString, _ctx);
        }

        protected override void EndProcessing()
        {
            _ctx.Writer.Flush();
            _ctx.Dispose();
            _dbUtil.Dispose();
        }
    }
}
