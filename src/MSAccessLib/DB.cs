﻿using Microsoft.Office.Interop.Access.Dao;
using Microsoft.VisualBasic;

namespace MSAccessLib
{
    public class DB : IDisposable
    {
        #region Other

        DBEngine _eng = null!;
        Workspace _wk = null!;

        public DB()
        {
            _eng = new DBEngineClass();
            _wk = _eng.Workspaces[0];
        }
        public void Dispose()
        {
            _wk?.Close();
            _wk = null;
            _eng = null;
        }

        #endregion

        #region Data Provider

        public Database? OpenExcel(string fileName) => 
            OpenDB(fileName, Connect.Excel(fileName));

        public Database? OpenAccessDB(string fileName, string pwd = "") => 
            OpenDB(fileName, Connect.MSAccess(fileName, pwd));

        public Database? OpenODBC(string connString, bool openreadonly = false) =>
            OpenDB("", connString, openreadonly);

        public Database? OpenDB(string name, string connectString, bool openreadonly = false) => 
            _wk.OpenDatabase(name, false, openreadonly, connectString);

        public Database? CreateMSAccess(string fileName) =>
            File.Exists(fileName) ? null :
                _wk.CreateDatabase(
                    fileName,
                    string.Format("{0}", LanguageConstants.dbLangGeneral),
                    DatabaseTypeEnum.dbVersion150
                );

        public Database? OpenDBWithConnString(string cnnString, string pwd = "") => cnnString switch
            {
                var c when c.StartsWith("ODBC;") => OpenODBC(c, true),
                var c when c.Contains(".xls") => OpenExcel(c),
                var c when c.Contains(".accdb") || c.Contains(".mdb") => OpenAccessDB(c, pwd),
                _ => OpenDB("", cnnString, true)
            };

        #endregion

        public void PrintDatabase(string cnnString, Context ctx)
        {
            var db = OpenDBWithConnString(cnnString);
            db?.PrintAll(ctx);
        }

        public void LinkTables(string srcCnnString, string destCnnString, Context ctx)
        {
            Database? srcDB = null;
            Database? destDB = null;

            try
            {
                destDB = OpenDBWithConnString(destCnnString);
                srcDB = OpenDBWithConnString(srcCnnString);
                destDB?.LinkToTables(srcDB!, ctx);
            }
            finally
            {
                srcDB?.Close();
                destDB?.Close();
                srcDB = null;
                destDB = null;
            }
        }

        public void ImportTables(string srcCnnString, string destCnnString, Context ctx)
        {
            Database? srcDB = null;
            Database? destDB = null;

            try
            {
                destDB = OpenDBWithConnString(destCnnString);
                srcDB = OpenDBWithConnString(srcCnnString);
                destDB?.ImportTables(srcDB!, ctx);
            }
            finally
            {
                srcDB?.Close();
                destDB?.Close();
                srcDB = null;
                destDB = null;
            }
        }
    }
}