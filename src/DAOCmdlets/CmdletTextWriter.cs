using System.Management.Automation;
using System.Text;

namespace DAOCmdlets
{
    public class CmdletTextWriter : TextWriter, IDisposable
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
                _cmd.WriteObject(_sb.ToString());
                _sb.Clear();
            }
        }

        public void Dispose()
        {
            base.Dispose();
            _cmd = null!;
        }
    }
}
