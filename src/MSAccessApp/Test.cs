using Microsoft.Office.Interop.Access.Dao;
using MSAccessLib;

namespace MSAccessApp
{
    public class Test
    {
        public void Run()
        {
            string f1 = @"c:\dev\Routing.accdb";
            string f2 = @"c:\dev\RoutingLink.accdb";
            string f3 = @"c:\dev\TestNewDB.accdb";
            //OpenMSAccess(f2);
            //OpenMSAccess(f1);
            //OpenExcel();
            //OpenODBC();
            //PrintAccessTables(f1);
            //CreateMSAccess(f3);
            //LinkTables(f1);
            ImportTables(f1);
        }

        protected void OpenMSAccess(string fileName)
        {
            using var t = new DB();
            Database? db = null;
            try
            {
                db = t.OpenAccessDB(fileName);
                db?.Print(true);
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
                db?.Print(true);
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
                db?.Print(true);
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
                db?.Print(true);
            }
            finally { db?.Close(); }
        }


        protected void LinkTables(string fileName)
        {
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
                destDB?.LinkToTables(srcDB!, t => t.Attributes == 0 || t.Attributes == 1073741824);
            }
            finally
            {
                destDB?.Close();
                srcDB?.Close();
            }
        }

        protected void ImportTables(string fileName)
        {
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
                destDB?.ImportTables(srcDB!, t => t.Attributes == 0 || t.Attributes == 1073741824);
            }
            finally
            {
                destDB?.Close();
                srcDB?.Close();
            }
        }
    }
}
