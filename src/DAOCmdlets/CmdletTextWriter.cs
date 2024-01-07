using MSAccessLib;
using System.Management.Automation;
using System.Text;

namespace DAOCmdlets
{
    /// <summary>
    /// This class wraps the cmdlet and redirects standard 
    /// TextWriter methods to standard Cmdlet WriteObject medthod.
    /// </summary>
    public class CmdletTextWriter : TextWriter, IDisposable2
    {
        Cmdlet _cmd = null!;
        StringBuilder _sb = null!;

        public CmdletTextWriter(Cmdlet cmd)
        {
            _cmd = cmd;
            _sb = new StringBuilder();
        }

        public override Encoding Encoding => Encoding.Default;

        public override void Write(char value)
        {
            // Cache each char until encounter a new line character
            // then write out the text as a line item similar to
            // WriteLine.
            if (value == '\r')
                return;
            else if (value == '\n' && _sb.Length > 0)
                CmdletWriteObject();
            else
                _sb.Append(value);
        }

        public override void Flush()
        {
            base.Flush();
            CmdletWriteObject();
        }

        protected void CmdletWriteObject()
        {
            if (_sb.Length > 0)
            {
                // write out the text to Powershell pipeline
                _cmd.WriteObject(_sb.ToString());
                _sb.Clear();
            }
        }

        void IDisposable2.Dispose()
        {
            base.Dispose();
            _cmd = null!;
        }
    }
}
