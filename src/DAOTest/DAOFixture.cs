using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using MSAccessApp;
using MSAccessLib;
using CF = MSAccessApp.AppConfigExtensions;

namespace DAOTest
{
    public class DAOFixture : IDisposable
    {
        DB _db = null!;
        IConfigurationRoot _configRoot = null!;
        Context _context = null!;
        IServiceProvider _provider = null!;

        public DAOFixture()
        {
            _db = new DB();
            _configRoot = CF.LoadConfig();
            _context = new Context();
            _provider = new ServiceCollection()
                .ConfigureServices(_configRoot)
                .BuildServiceProvider();
        }

        public DB DB => _db;

        public (string DSN, string PWD) OdbcDsn
        {
            get
            {
                var cnnstr = _configRoot.GetConnectionString("OdbcDsn")!.Split(";");
                if (cnnstr.Length > 1)
                    return (cnnstr[0], cnnstr[1]);
                else return (cnnstr[0], "");
            }
        }

        public (string FilePath, string PWD) MSAccess
        {
            get
            {
                var cnnstr = _configRoot.GetConnectionString("MSAccess")!.Split(";");
                if (cnnstr.Length > 1)
                    return (cnnstr[0], cnnstr[1]);
                else return (cnnstr[0], "");
            }
        }

        public string OdbcDsnless => _configRoot.GetConnectionString("OdbcDsnless2")!;

        public string Excel => _configRoot.GetConnectionString("Excel")!;

        public (string FileName, string Pwd) OdbcDsnFile 
        {
            get 
            {
                var cnnstr = _configRoot.GetConnectionString("OdbcDsnFile")!.Split(";");
                if (cnnstr.Length > 1)
                    return (cnnstr[0], cnnstr[1]);
                else return (cnnstr[0], "");
            }
        }

        public Context Context => _context;

        public void Dispose()
        {
            _db.Dispose();
        }
    }
}
