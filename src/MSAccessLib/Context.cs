using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Office.Interop.Access.Dao;

namespace MSAccessLib
{
    public record Context : IDisposable
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
            GetNewTableName = t => t.Name;
            IsSavePwdWithLinkTable = false;
        }

        public Predicate<TableDef> TableFilter { get; set; } = null!;
        public Predicate<QueryDef> QueryFilter { get; set; } = null!;
        public bool HideEmptyProperty { get; set; }
        public TextWriter Writer { get; set; } = null!;
        public ILogger Logger { get; set; } = null!;
        public bool HideFieldProperty { get; set; }
        public bool IsMSAccessDB { get; set; }
        public Func<TableDef, string> GetNewTableName { get; set; }
        public bool IsSavePwdWithLinkTable { get; set; }

        public void Dispose()
        {
            if (Logger is IDisposable2 l)
                l.Dispose();
            Logger = null!;

            if (Writer is IDisposable2 w)
                w.Dispose();
            Writer = null!;
        }
    }
}
