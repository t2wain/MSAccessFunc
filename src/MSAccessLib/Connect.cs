using Microsoft.Office.Interop.Access.Dao;
using System.Text;
using System.Xml.Linq;

namespace MSAccessLib
{
    public static class Connect
    {
        public static string DSNFile(string dsnFileName)
        {
            return string.Format("ODBC;FILEDSN={0}", dsnFileName);
        }

        public static string DSN(string dsnName, string pwd = "")
        {
            if (string.IsNullOrWhiteSpace(pwd))
                return string.Format("ODBC;DSN={0};", dsnName);
            else return string.Format("ODBC;DSN={0};Pwd={1}", dsnName, pwd);
        }

        public static string Oracle(string server, string uid, string pwd)
        {
            // Driver={Microsoft ODBC for Oracle};Server=myServerAddress;Uid=myUsername;Pwd=myPassword;

            // Driver={Microsoft ODBC for Oracle};
            // Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=199.199.199.199)(PORT=1523))(CONNECT_DATA=(SID=dbName)));
            // Uid=myUsername;Pwd=myPassword;

            // Driver={Microsoft ODBC for Oracle};
            // CONNECTSTRING=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=server)(PORT=7001))(CONNECT_DATA=(SERVICE_NAME=myDb)));
            // Uid=myUsername;Pwd=myPassword;

            return "ODBC;Driver={Microsoft ODBC for Oracle};" + string.Format("Server={0};Uid={1};Pwd={2};",
                server, uid, pwd);
        }

        public static string SqlServer(string server, string dbName)
        {
            // Driver={SQL Server};Server=myServerAddress;Database=myDataBase;Trusted_Connection=Yes;

            return "ODBC;Driver={SQL Server};" + string.Format("Server={0};Database={1};Trusted_Connection=Yes;",
                server, dbName);
        }

        public static string Sqlite(string dbFileName)
        {
            // DRIVER=SQLite3 ODBC Driver;Database=c:\mydb.db;LongNames=0;Timeout=1000;NoTXN=0;SyncPragma=NORMAL;StepAPI=0;

            return "ODBC;DRIVER=SQLite3 ODBC Driver;" + 
                string.Format("Database={0};LongNames=0;Timeout=1000;NoTXN=0;SyncPragma=NORMAL;StepAPI=0;",
                    dbFileName); 
        }

        public static string MSAccess(string filename, string pwd)
        {
            var c = new StringBuilder();
            c.Append(string.Format("MS Access;DATABASE={0}", filename));
            if (!string.IsNullOrWhiteSpace(pwd))
                c.Append(string.Format(";PWD={0}", pwd));
            return c.ToString();
        }

        public static string Excel(string filename)
        {
            var c = new StringBuilder();
            c.Append(filename.EndsWith(".xlsx") ? "Excel 12.0;HDR=Yes;" : "Excel 8.0;HDR=Yes;");
            if (!string.IsNullOrWhiteSpace(filename))
                c.Append(string.Format("DATABASE={0};", filename));
            return c.ToString();
        }

        public static string GetConnectString(this Database db)
        {
            var cnnstr = "";
            if (db.Name.EndsWith(".accdb") || db.Name.EndsWith(".mdb"))
            {
                cnnstr = db.Connect;
                cnnstr = string.IsNullOrWhiteSpace(cnnstr) ? Connect.MSAccess(db.Name, "") : cnnstr;
                cnnstr = cnnstr.StartsWith("MS Access") ? cnnstr : "MS Access" + cnnstr;
            }
            else cnnstr = db.Connect;

            return cnnstr;
        }

    }
}
