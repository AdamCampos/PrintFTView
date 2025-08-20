using System;
using System.IO;
using System.Text;

namespace LibFTView.Services
{
    internal static class DiagLog
    {
        private static readonly object _lock = new object();
        private static string _logDir = @"C:\Projetos\VisualStudio\LibFTView";
        private static string _logFile = "net.log";

        public static void SetPath(string folder, string fileName = "net.log")
        {
            if (!string.IsNullOrWhiteSpace(folder)) _logDir = folder;
            if (!string.IsNullOrWhiteSpace(fileName)) _logFile = fileName;
            try { Directory.CreateDirectory(_logDir); } catch { /* ignore */ }
        }

        public static void Write(string msg)
        {
            try
            {
                lock (_lock)
                {
                    Directory.CreateDirectory(_logDir);
                    var path = Path.Combine(_logDir, _logFile);
                    var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}  {msg}";
                    File.AppendAllText(path, line + Environment.NewLine, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
                }
            }
            catch { /* não quebra o fluxo */ }
        }
    }
}
