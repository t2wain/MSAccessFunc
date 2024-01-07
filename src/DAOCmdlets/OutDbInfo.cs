﻿using MSAccessLib;
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
        public ScriptBlock TableFilter { get; set; } = ScriptBlock.Create("$true");

        [Parameter]
        public ScriptBlock QueryFilter { get; set; } = ScriptBlock.Create("$true");

        protected override void BeginProcessing()
        {
            _dbUtil = new DB();
            _ctx = this.BuildContext(ConnectString, TableFilter, QueryFilter);
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