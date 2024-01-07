using Microsoft.Extensions.Logging;
using MSAccessLib;
using System.Management.Automation;

namespace DAOCmdlets
{
    /// <summary>
    /// This class wraps the cmdlet and redirects standard 
    /// ILogger log methods to standard Cmdlet Write{XXX} methods.
    /// </summary>
    public class CmdletLogger : ILogger, IDisposable2
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

        void IDisposable2.Dispose()
        {
            _cmd = null!;
        }

    }
}
