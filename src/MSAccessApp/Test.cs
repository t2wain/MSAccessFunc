using Microsoft.Extensions.Logging;
using Microsoft.Office.Interop.Access.Dao;
using MSAccessLib;

namespace MSAccessApp
{
    public class Test
    {
        Predicate<TableDef> _filter = null!;
        ILoggerFactory _factory = null!;

        public Test()
        {
            _factory = LoggerFactory.Create(builder => {
                builder.AddConsole();
                //builder.SetMinimumLevel(LogLevel.None);
            });
            _filter = t => t.Attributes == 0 || t.Attributes == 1073741824;
        }

        public void Run()
        {
            string f1 = @"c:\dev\Routing.accdb";
            string f2 = @"c:\dev\RoutingLink.accdb";
            string f3 = @"c:\dev\TestNewDB.accdb";
            string f4 = @"c:\dev\RoutingInfo.txt";
            //OpenMSAccess(f2);
            //OpenMSAccess(f1);
            //OpenExcel();
            //OpenODBC();
            PrintAccessTables(f1);
            //PrintAccessTables2(f1, f4);
            //CreateMSAccess(f3);
            //LinkTables(f1);
            //ImportTables(f1);
        }

        protected void OpenMSAccess(string fileName)
        {
            using var t = new DB();
            Database? db = null;
            try
            {
                db = t.OpenAccessDB(fileName);
                db?.Print(Console.Out, true);
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
                db?.Print(Console.Out, true);
            }
            finally { db?.Close(); }
        }

        protected void OpenODBC()
        {
            using var t = new DB();
            Database? db = null;
            try
            {
                db = t.OpenODBC(@"RoutingDB");
                db?.Print(Console.Out, true);
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
            try
            {
                db = t.OpenAccessDB(fileName);
                db?.Print(Console.Out, true);
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
            try
            {
                db = t.OpenAccessDB(dbFileName);
                db?.Print(wr, true, _filter, logger);
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
                destDB?.LinkToTables(srcDB!, _filter, logger);
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
                destDB?.ImportTables(srcDB!, _filter, logger);
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
