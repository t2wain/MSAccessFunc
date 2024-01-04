using Microsoft.Office.Interop.Access.Dao;
using System.Text;
using System.Xml.Linq;

namespace MSAccessLib
{
    public static class Connect
    {
        /// <summary>
        /// Note: you can use the ODBC Data Source Administrator app to build the DSN file.
        /// On the "File DSN" tab, click (1) Add, (2) select an ODBC driver, (3) click Browse to pick
        /// a folder to save the DSN file, and (4) next a custom windows for the selected
        /// ODBC driver will appear to prompt you for the DB parameters
        /// </summary>
        public static string DSNFile(string dsnFileName, string pwd = "")
        {

            if (string.IsNullOrWhiteSpace(pwd))
                return string.Format("ODBC;FILEDSN={0};", dsnFileName);
            else return string.Format("ODBC;FILEDSN={0};Pwd={1}", dsnFileName, pwd);
        }

        public static string DSN(string dsnName, string pwd = "")
        {
            if (string.IsNullOrWhiteSpace(pwd))
                return string.Format("ODBC;DSN={0};", dsnName);
            else return string.Format("ODBC;DSN={0};Pwd={1}", dsnName, pwd);
        }

        /// <summary>
        /// Note, the connect string format is different for each ODBC driver. 
        /// </summary>
        public static string Oracle(string odbcDriver, string tnsname, string uid, string pwd)
        {
            // Driver={Microsoft ODBC for Oracle};Server=myServerAddress;Uid=myUsername;Pwd=myPassword;

            // Driver={Microsoft ODBC for Oracle};
            // Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=199.199.199.199)(PORT=1523))(CONNECT_DATA=(SID=dbName)));
            // Uid=myUsername;Pwd=myPassword;

            // Driver={Microsoft ODBC for Oracle};
            // CONNECTSTRING=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=server)(PORT=7001))(CONNECT_DATA=(SERVICE_NAME=myDb)));
            // Uid=myUsername;Pwd=myPassword;

            // Driver={Oracle in instantclient_12_1};
            // Dbq=tnsname;
            // Uid=myUsername;Pwd=myPassword;

            // DRIVER={Oracle in instantclient_12_2};DBQ=HOST:PORT/SERVICE_NAME;UID=USERNAME;PWD=PASSWORD
            // DRIVER={Oracle in instantclient_12_2};DBQ=127.0.0.1:1521/ORCLPDB1.localdomain;UID=sys;PWD=Oradoc_db1 as sysdba
            // Note: SERVICE_NAME is the same as SID

            return string.Format("ODBC;Driver={{{0}}};Dbq={1};Uid={2};Pwd={3};",
                odbcDriver, tnsname, uid, pwd);
        }

        /// <summary>
        /// Note, the connect string format is different for each ODBC driver. 
        /// </summary>
        public static string Oracle(string odbcDriver, string host, int port, string sid, string uid, string pwd)
        {
            if (odbcDriver.Contains("Microsoft"))
                return string.Format(
                    "ODBC;Driver={{{0}}};" + 
                    "Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={1})(PORT={2}))(CONNECT_DATA=(SID={3})));" +
                    "Uid={4};Pwd={5};",
                    odbcDriver, host, port, sid, uid, pwd);
            else
                return string.Format(
                    "ODBC;Driver={{{0}}};" +
                    "DBQ={1}:{2}/{3};" +
                    "Uid={4};Pwd={5};",
                    odbcDriver, host, port, sid, uid, pwd);
        }

        public static string SqlServer(string server, string dbName)
        {
            // Driver={SQL Server};Server=myServerAddress;Database=myDataBase;Trusted_Connection=Yes;

            return string.Format("ODBC;Driver={{SQL Server}};Server={0};Database={1};Trusted_Connection=Yes;",
                server, dbName);
        }

        public static string Sqlite(string dbFileName)
        {
            // DRIVER=SQLite3 ODBC Driver;Database=c:\mydb.db;LongNames=0;Timeout=1000;NoTXN=0;SyncPragma=NORMAL;StepAPI=0;

            return "ODBC;DRIVER=SQLite3 ODBC Driver;" + 
                string.Format("Database={0};LongNames=0;Timeout=1000;NoTXN=0;SyncPragma=NORMAL;StepAPI=0;",
                    dbFileName); 
        }

        public static string MSAccess(string filename, string pwd = "")
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
