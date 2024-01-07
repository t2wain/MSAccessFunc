using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace DAOCmdlets
{
    public class CmdletLogger : ILogger, IDisposable
    {
        Cmdlet _cmd = null!;

        public CmdletLogger(Cmdlet cmd)
        {
            _cmd = cmd;
        }
        public LogLevel LogLevel { get; set; } = LogLevel.Information;

        public IDisposable? BeginScope<TState>(TState state) => default!;

        public bool IsEnabled(LogLevel logLevel) => true;

        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state,
            Exception? exception, Func<TState, Exception?, string> formatter)
        {
            var msg = formatter(state, exception);
            switch (logLevel)
            {
                case LogLevel.Error:
                    _cmd.WriteError(new ErrorRecord(exception, eventId.Name, ErrorCategory.OperationStopped, null));
                    break;
                case LogLevel.Debug:
                    _cmd.WriteDebug(msg);
                    break;
                case LogLevel.Warning:
                    _cmd.WriteWarning(msg);
                    break;
                case LogLevel.Trace:
                    _cmd.WriteVerbose(msg);
                    break;
                default:
                    _cmd.WriteInformation(new InformationRecord(msg, ""));
                    break;
            }
        }

        public void Dispose()
        {
            _cmd = null!;
        }

    }
}
