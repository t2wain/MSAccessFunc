using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Office.Interop.Access.Dao;
using MSAccessLib;

namespace MSAccessApp
{
    public class Examples
    {
        private readonly Context _ctx = null!;
        private readonly IServiceProvider _provider;
        private readonly IConfigurationSection _cfg;

        public Examples(IServiceProvider provider)
        {
            _provider = provider;
            _cfg = _provider.GetRequiredService<IConfigurationRoot>().GetSection("TestParams");
            _ctx = new Context()
            {
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
            int testNo = 9;

            var dsnName = _cfg["OdbcDsn"]!;
            var dsnLinkDbFile = _cfg["OdbcMSAccessLinkFile"]!;
            var dbFile = _cfg["MSAccessFile"]!;
            var dbLinkFile = _cfg["MSAccessLinkFile"]!;
            var dbImportFile = _cfg["MSAccessImportFile"]!;

            switch (testNo)
            {
                case 0:
                    // print db info of an ODBC db (using DSN) info to console
                    PrintDSN(dsnName);
                    break;
                case 1:
                    // print db info of an ODBC db (using DSN) info to file
                    var dsnInfoFile = _cfg["OdbcDsnInfoFile"]!;
                    PrintDSN(dsnName, dsnInfoFile);
                    break;
                case 2:
                    // link tables of an ODBC db (using DSN) to an Access file
                    LinkDsnTables(dsnName, dsnLinkDbFile);
                    break;
                case 3:
                    // print db info to console of an Access db
                    // having linked table from an ODBC db
                    PrintAccess(dsnLinkDbFile);
                    break;
                case 4:
                    // print db info to a file of an Access file
                    // having linked table from an ODBC db
                    var dsnLinkInfoFile = _cfg["OdbcMSAccessLinkInfoFile"]!;
                    PrintAccess(dsnLinkDbFile, dsnLinkInfoFile);
                    break;
                case 5:
                    // print db info of an Access db to console
                    PrintAccess(dbFile);
                    break;
                case 6:
                    // print db info of an Access db to file
                    var dbFileInfo = _cfg.GetValue<string>("MSAccessInfoFile")!;
                    PrintAccess(dbFile, dbFileInfo);
                    break;
                case 7:
                    // link tables of an Access db to another Access db
                    LinkAccessTables(dbFile, dbLinkFile);
                    break;
                case 8:
                    // print db info of an Access db with linked tables to console
                    PrintAccess(dbLinkFile);
                    break;
                case 9:
                    // import tables from an Access db to another Access db
                    ImportAccess(dbFile, dbImportFile);
                    break;
            }
        }

        #region Access DB

        protected void PrintAccess(string dbFile, string dbInfoFile = "")
        {
            using var t = new DB();
            Database? db = null;
            var ctx = _ctx with { HideFieldProperty = true };
            try
            {
                db = t.OpenAccessDB(dbFile);
                PrintDB(db!, ctx, dbInfoFile);
            }
            finally { db?.Close(); }
        }

        protected void LinkAccessTables(string dbFile, string linkDbFile)
        {
            using var t = new DB();
            if (File.Exists(linkDbFile))
                File.Delete(linkDbFile);
            Database? destDB = null;
            Database? srcDB = null;
            try
            {
                destDB = t.CreateMSAccess(linkDbFile);
                srcDB = t.OpenAccessDB(dbFile);
                var ctx = _ctx with 
                { 
                    Logger = _provider.GetRequiredService<ILoggerFactory>().CreateLogger("Linking") 
                };
                destDB?.LinkToTables(srcDB!, ctx);
            }
            finally
            {
                destDB?.Close();
                srcDB?.Close();
            }
        }

        protected void ImportAccess(string dbFile, string dbImportFile)
        {
            var logger = _provider.GetRequiredService<ILoggerFactory>().CreateLogger("Importing");
            var ctx = _ctx with { Logger = logger };

            using var t = new DB();
            Database? srcDB = null;
            try
            {
                srcDB = t.OpenAccessDB(dbFile);
                ImportTable(srcDB!, ctx, dbImportFile, t);
            }
            finally
            {
                srcDB?.Close();
            }
        }

        #endregion

        #region DSN

        protected void PrintDSN(string dsnName, string dsnInfoFile = "")
        {
            using var t = new DB();
            Database? db = null;

            try
            {
                db = t.OpenODBC(Connect.DSN(dsnName));
                var ctx = _ctx with
                {
                    TableFilter = t => t.Name.StartsWith("VW.C_SW_"),
                    QueryFilter = q => q.Name.StartsWith("VW.C_VW_SW_"),
                    IsMSAccessDB = false
                };
                PrintDB(db!, ctx, dsnInfoFile);
            }
            finally { db?.Close(); }
        }

        protected void LinkDsnTables(string dsnName, string dsnLinkDbFile)
        {
            var ctx = _ctx with
            {
                Logger = _provider.GetRequiredService<ILoggerFactory>().CreateLogger("DSN Linking"),
                TableFilter = t => t.Name.StartsWith("VW.C_"),
                GetLinkTableName = t => t.Name.Replace("VW.", ""),
                IsSavePwdWithLinkTable = true
            };

            using var t = new DB();
            if (File.Exists(dsnLinkDbFile))
                File.Delete(dsnLinkDbFile);
            Database? destDB = null;
            Database? srcDB = null;
            try
            {
                destDB = t.CreateMSAccess(dsnLinkDbFile);
                srcDB = t.OpenODBC(Connect.DSN(dsnName));
                destDB?.LinkToTables(srcDB!, ctx);
            }
            finally
            {
                destDB?.Close();
                srcDB?.Close();
            }
        }

        #endregion

        #region Common

        protected void ImportTable(Database srcDb, Context ctx, string dbImportFile, DB t)
        {
            if (File.Exists(dbImportFile))
                File.Delete(dbImportFile);
            Database? destDB = null;
            try
            {
                destDB = t.CreateMSAccess(dbImportFile);
                destDB?.ImportTables(srcDb, ctx);
            }
            finally
            {
                destDB?.Close();
            }
        }

        protected void PrintDB(Database db, Context ctx, string dbInfoFile = "")
        {
            if (!string.IsNullOrEmpty(dbInfoFile))
            {
                if (File.Exists(dbInfoFile))
                    File.Delete(dbInfoFile);
                using var f = File.Open(dbInfoFile, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.Read);
                using var wr = new StreamWriter(f);
                ctx = ctx with
                {
                    Writer = wr,
                    Logger = _provider.GetRequiredService<ILoggerFactory>().CreateLogger("PrintDB")
                };
                db?.PrintAll(ctx);
            }
            else
            {
                db?.PrintAll(ctx);
            }
        }

        #endregion
    }
}
