using System;
using System.IO;
using System.Text;

namespace Smart3D.ExcelImport.Services
{
    /// <summary>
    /// Logging service for the import operation.
    /// Provides structured logging to file and/or console.
    /// </summary>
    public sealed class LoggingService : IDisposable
    {
        private readonly StreamWriter _logWriter;
        private readonly string _logFilePath;
        private readonly object _lock = new object();
        private bool _disposed;

        public event EventHandler<LogEventArgs> LogEntry;

        public LoggingService(string logFilePath = null)
        {
            _logFilePath = logFilePath ?? Path.Combine(
                Path.GetTempPath(), 
                $"Smart3DImport_{DateTime.Now:yyyyMMdd_HHmmss}.log");

            _logWriter = new StreamWriter(_logFilePath, true, Encoding.UTF8)
            {
                AutoFlush = true
            };

            LogInfo($"Log initialized at {_logFilePath}");
        }

        public void LogInfo(string message)
        {
            WriteLog("INFO", message);
        }

        public void LogWarning(string message)
        {
            WriteLog("WARN", message);
        }

        public void LogError(string message, Exception ex = null)
        {
            var fullMessage = ex != null 
                ? $"{message}\nException: {ex.GetType().Name}: {ex.Message}\nStackTrace: {ex.StackTrace}"
                : message;
            WriteLog("ERROR", fullMessage);
        }

        public void LogDebug(string message)
        {
            WriteLog("DEBUG", message);
        }

        private void WriteLog(string level, string message)
        {
            var logLine = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {message}";

            lock (_lock)
            {
                try
                {
                    _logWriter?.WriteLine(logLine);
                }
                catch { /* Ignore logging errors */ }
            }

            LogEntry?.Invoke(this, new LogEventArgs { Level = level, Message = message });
        }

        public string GetLogFilePath() => _logFilePath;

        public string GetFullLog()
        {
            lock (_lock)
            {
                _logWriter?.Flush();
                if (File.Exists(_logFilePath))
                {
                    try { return File.ReadAllText(_logFilePath); }
                    catch { return "Unable to read log file."; }
                }
                return "No log file available.";
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                lock (_lock)
                {
                    _logWriter?.Flush();
                    _logWriter?.Dispose();
                }
                _disposed = true;
            }
        }
    }

    public class LogEventArgs : EventArgs
    {
        public string Level { get; set; }
        public string Message { get; set; }
    }
}
