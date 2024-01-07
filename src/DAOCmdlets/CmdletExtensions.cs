using MSAccessLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace DAOCmdlets
{
    public static class CmdLetExtensions
    {
        public static Context BuildContext(this Cmdlet cmd, string cnnstring, ScriptBlock tableFilter, ScriptBlock queryFilter)
        {
            return new Context()
            {
                TableFilter = t => {
                    foreach (var r in tableFilter.Invoke(t))
                    {
                        var b = (bool)r.BaseObject;
                        return b;
                    }
                    return false;
                },
                QueryFilter = q => {
                    foreach (var r in queryFilter.Invoke(q))
                    {
                        var b = (bool)r.BaseObject;
                        return b;
                    }
                    return false;
                },
                Writer = new CmdletTextWriter(cmd),
                Logger = new CmdletLogger(cmd),
                IsMSAccessDB =
                    cnnstring.Contains(".accdb")
                    || cnnstring.Contains(".mdb")

            };
        }
    }
}
