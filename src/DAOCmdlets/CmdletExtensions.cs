using Microsoft.Office.Interop.Access.Dao;
using MSAccessLib;
using System.Management.Automation;

namespace DAOCmdlets
{
    public static class CmdLetExtensions
    {
        public static Context BuildContext(this Cmdlet cmd, ScriptBlock? tableFilter = null, 
            ScriptBlock? queryFilter = null, ScriptBlock? getDestTableName = null)
        {
            Predicate<TableDef> tf = tableFilter switch
            {
                null => t => true,
                _ => t =>
                    {
                        // redirect a .net predicate to a Powershell scriptblock
                        foreach (var r in tableFilter.Invoke(t))
                        {
                            var b = (bool)r.BaseObject;
                            return b;
                        }
                        return false;
                    }
            };

            Predicate<QueryDef> qf = queryFilter switch
            {
                null => q => true,
                _ => q =>
                {
                    // redirect a .net predicate to a Powershell scriptblock
                    foreach (var r in queryFilter.Invoke(q))
                    {
                        var b = (bool)r.BaseObject;
                        return b;
                    }
                    return false;
                }
            };

            Func<TableDef, string> gn = getDestTableName switch
                {
                    null => t => t.Name,
                    _ => t =>
                        {
                            foreach (var r in getDestTableName.Invoke(t))
                            {
                                var n = (string)r.BaseObject;
                                return n;
                            }
                            return t.Name;
                        }
                };

            return new Context()
            {
                TableFilter = tf,
                QueryFilter = qf,
                GetDestTableName = gn,
                // redirect TextWriter methods to Cmdlet WriteObject method
                Writer = new CmdletTextWriter(cmd),
                // redirect ILogger methods to Cmdlet Write{XXX} methods
                Logger = new CmdletLogger(cmd),
            };
        }
    }
}
