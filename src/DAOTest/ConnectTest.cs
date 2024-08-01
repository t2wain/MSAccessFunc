using Microsoft.Office.Interop.Access.Dao;
using MSAccessLib;

namespace DAOTest
{
    public class ConnectTest : IClassFixture<DAOFixture>
    {
        DAOFixture _fxt = null!;

        public ConnectTest(DAOFixture fxt)
        {
            _fxt = fxt;
        }

        [Fact]
        public void Should_open_MSAccess()
        {
            Database? db = null;
            try {
                var cnnstr = _fxt.MSAccess;
                db = _fxt.DB.OpenAccessDB(cnnstr.FilePath, cnnstr.PWD);
                var res = db?.IsMsAccess();
                Assert.NotNull(db);
            }
            finally { db?.Close(); }
        }

        [Fact]
        public void Should_open_Excel()
        {
            Database? db = null;
            try
            {
                var cnnstr = _fxt.Excel;
                db = _fxt.DB.OpenExcel(cnnstr);
                var res = db?.IsMsAccess();
                Assert.NotNull(db);
            }
            finally { db?.Close(); }
        }

        [Fact]
        public void Should_open_ODBC_DSN()
        {
            Database? db = null;
            try
            {
                var cnnstr = _fxt.OdbcDsn;
                db = _fxt.DB.OpenODBC(Connect.DSN(cnnstr.DSN, cnnstr.PWD));
                var res = db?.IsMsAccess();
                Assert.NotNull(db);
            }
            finally { db?.Close(); }
        }

        [Fact]
        public void Should_open_ODBC_FileDSN()
        {
            Database? db = null;
            try
            {
                var (filename, pwd) = _fxt.OdbcDsnFile;
                db = _fxt.DB.OpenODBC(Connect.DSNFile(filename, pwd));
                Assert.NotNull(db);
            }
            finally { db?.Close(); }
        }

        [Fact]
        public void Should_open_ODBC_DSNless()
        {
            Database? db = null;
            try
            {
                var cnnstr = _fxt.OdbcDsnless;
                db = _fxt.DB.OpenODBC(cnnstr);
                Assert.NotNull(db);
            }
            finally { db?.Close(); }
        }
    }
}