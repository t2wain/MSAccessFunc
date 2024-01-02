using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Office.Interop.Access.Dao;
using MSAccessLib;

namespace MSAccessApp
{
    public class Test
    {
        ILoggerFactory _factory = null!;
        Context _ctx = null!;

        public Test()
        {
            _factory = LoggerFactory.Create(builder => {
                builder.AddConsole();
                //builder.SetMinimumLevel(LogLevel.None);
            });
            _ctx = new Context()
            {
                //TableFilter = t => new int[] { 0, 1073741824, 536870912, 537001984 }.Contains(t.Attributes),
                TableFilter = t => !new int[] { 
                    (int)TableDefAttributeEnum.dbHiddenObject,
                    (int)TableDefAttributeEnum.dbSystemObject,
                    -2147483648,
                    2
                }.Contains(t.Attributes),
                HideEmptyProperty = true,
                HideFieldProperty = true,
                IsMSAccessDB = true,
            };
        }

        public void Run()
        {
            var f1 = @"c:\dev\Routing.accdb";
            var f2 = @"c:\dev\RoutingLink.accdb";
            var f3 = @"c:\dev\TestNewDB.accdb";
            var f4 = @"c:\dev\RoutingInfo.txt";
            var f5 = @"c:\dev\TestLinkDB.accdb";
            var f6 = @"c:\dev\TestOdbcLink.accdb";
            var f7 = @"c:\dev\TestLinkDBInfo.txt";
            //OpenMSAccess(f2);
            //OpenMSAccess(f1);
            //OpenExcel();
            //OpenDSN(@"c:\dev\KBRSoftwareDBInfo.txt");
            //PrintAccessTables(f6);
            //PrintAccessTables2(f5, f7);
            //CreateMSAccess(f3);
            //LinkTables(f1);
            //ImportTables(f1);
            LinkDsnTables();
        }

        protected void OpenMSAccess(string fileName)
        {
            using var t = new DB();
            Database? db = null;
            try
            {
                db = t.OpenAccessDB(fileName);
                db?.PrintAll(_ctx);
            }
            finally { db?.Close(); }
        }

        protected void OpenExcel()
        {
            using var t = new DB();
            Database? db = null;
            try
            {
                db = t.OpenExcel(@"c:\dev\SampleTable.xlsx");
                db?.PrintAll(_ctx);
            }
            finally { db?.Close(); }
        }

        protected void OpenDSN(string outFileName)
        {
            var logger = _factory.CreateLogger("KBRSoftware");
            using var t = new DB();
            Database? db = null;
            if (File.Exists(outFileName))
                File.Delete(outFileName);
            using var f = File.Open(outFileName, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.Read);
            using var wr = new StreamWriter(f);
            var ctx = new Context()
            {
                TableFilter = t => t.Name.StartsWith("VW.C_SW_"),
                QueryFilter = q => q.Name.StartsWith("VW.C_VW_SW_"),
                Writer = wr,
                Logger = logger,
                HideEmptyProperty = true,
                HideFieldProperty = true,
                IsMSAccessDB = false
            };
            try
            {
                db = t.OpenODBC(Connect.DSN("KBRSofware PROD"));
                db?.PrintAll(ctx);
            }
            finally { db?.Close(); }
        }

        protected void CreateMSAccess(string fileName)
        {
            using var t = new DB();
            Database? db = null;
            try
            {
                db = t.CreateMSAccess(fileName);
            }
            finally { db?.Close(); }
        }

        protected void PrintAccessTables(string fileName)
        {
            using var t = new DB();
            Database? db = null;
            var ctx = _ctx with { HideFieldProperty = true };
            try
            {
                db = t.OpenAccessDB(fileName);
                db?.PrintAll(ctx);
            }
            finally { db?.Close(); }
        }

        protected void PrintAccessTables2(string dbFileName, string outFileName)
        {
            var logger = _factory.CreateLogger("DbInfo");
            using var t = new DB();
            Database? db = null;
            using var f = File.OpenWrite(outFileName);
            using var wr = new StreamWriter(f);
            var ctx = _ctx with { Logger = logger, Writer = wr, };
            try
            {
                db = t.OpenAccessDB(dbFileName);
                db?.PrintAll(ctx);
            }
            finally { 
                db?.Close();
                wr.Flush();
                wr.Close();
                f.Close();
            }
        }

        protected void LinkTables(string fileName)
        {
            var logger = _factory.CreateLogger("Linking");
            var ctx = _ctx with { Logger = logger };

            using var t = new DB();
            string testDb = @"c:\dev\TestLinkDB.accdb";
            if (File.Exists(testDb))
                File.Delete(testDb);
            Database? destDB = null;
            Database? srcDB = null;
            try
            {
                destDB = t.CreateMSAccess(testDb);
                srcDB = t.OpenAccessDB(fileName);
                destDB?.LinkToTables(srcDB!, ctx);
            }
            finally
            {
                destDB?.Close();
                srcDB?.Close();
            }
        }

        protected void LinkDsnTables()
        {
            var logger = _factory.CreateLogger("Linking");
            var ctx = _ctx with 
            { 
                Logger = logger,
                TableFilter = t => t.Name.StartsWith("VW.C_"),
                GetLinkTableName = t => t.Name.Replace("VW.", "") 
            };

            using var t = new DB();
            string testDb = @"c:\dev\TestLinkDB.accdb";
            if (File.Exists(testDb))
                File.Delete(testDb);
            Database? destDB = null;
            Database? srcDB = null;
            try
            {
                destDB = t.CreateMSAccess(testDb);
                srcDB = t.OpenODBC(Connect.DSN("KBRSofware PROD"));
                destDB?.LinkToTables(srcDB!, ctx);
            }
            finally
            {
                destDB?.Close();
                srcDB?.Close();
            }
        }

        protected void ImportTables(string fileName)
        {
            var logger = _factory.CreateLogger("Importing");
            var ctx = _ctx with { Logger = logger };

            using var t = new DB();
            string testDb = @"c:\dev\TestImportDB.accdb";
            if (File.Exists(testDb))
                File.Delete(testDb);
            Database? destDB = null;
            Database? srcDB = null;
            try
            {
                destDB = t.CreateMSAccess(testDb);
                srcDB = t.OpenAccessDB(fileName);
                destDB?.ImportTables(srcDB!, ctx);
            }
            finally
            {
                destDB?.Close();
                srcDB?.Close();
                _factory?.Dispose();
            }
        }
    }
}
