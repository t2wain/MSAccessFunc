using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Office.Interop.Access.Dao;

namespace MSAccessLib
{
    public record Context
    {
        public Context()
        {
            TableFilter = t => true;
            QueryFilter = t => true;
            HideEmptyProperty = false;
            Writer = Console.Out;
            Logger = NullLogger.Instance;
            HideFieldProperty = false;
            IsMSAccessDB = false;
            GetLinkTableName = t => t.Name;
            IsSavePwdWithLinkTable = false;
        }

        public Predicate<TableDef> TableFilter { get; set; } = null!;
        public Predicate<QueryDef> QueryFilter { get; set; } = null!;
        public bool HideEmptyProperty { get; set; }
        public TextWriter Writer { get; set; } = null!;
        public ILogger Logger { get; set; } = null!;
        public bool HideFieldProperty { get; set; }
        public bool IsMSAccessDB { get; set; }
        public Func<TableDef, string> GetLinkTableName { get; set; }
        public bool IsSavePwdWithLinkTable { get; set; }
    }
}
