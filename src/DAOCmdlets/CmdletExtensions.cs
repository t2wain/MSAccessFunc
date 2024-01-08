using Microsoft.Office.Interop.Access.Dao;
using MSAccessLib;
using System.Management.Automation;

namespace DAOCmdlets
{
    public static class CmdLetExtensions
    {
        /// <summary>
        /// A convenience method (1) to wrap the PowerShell ScriptBlock objects inside 
        /// the .NET Predicate functions of the Context object, and (2) to wrap the
        /// given cmdlet inside the CmdletLogger and CmdletTextWriter objects 
        /// that implement ILogger and TextWriter interfaces.
        /// </summary>
        /// <param name="cmd">The Cmdlet object</param>
        /// <param name="tableFilter">Given a TableDef return true for selected TableDef</param>
        /// <param name="queryFilter">Given a QueryDef return true for selected QueryDef</param>
        /// <param name="getDestTableName">Given a TableDef return a new name for TableDef</param>
        /// <returns></returns>
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
