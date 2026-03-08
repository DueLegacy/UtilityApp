using System;
using System.IO;
using UtilityApp.Contracts;

namespace UtilityApp
{
    internal sealed class HostContext : IHostContext
    {
        private readonly Action<string> _logAction;

        public HostContext(string applicationRootPath, Action<string> logAction)
        {
            ApplicationRootPath = applicationRootPath;
            ConfigDirectoryPath = Path.Combine(ApplicationRootPath, "Config");
            LogsDirectoryPath = Path.Combine(ApplicationRootPath, "Logs");
            _logAction = logAction;

            Directory.CreateDirectory(ConfigDirectoryPath);
            Directory.CreateDirectory(LogsDirectoryPath);
            Directory.CreateDirectory(Path.Combine(ApplicationRootPath, "Modules"));
        }

        public string ApplicationRootPath { get; private set; }

        public string ConfigDirectoryPath { get; private set; }

        public string LogsDirectoryPath { get; private set; }

        public void Log(string message)
        {
            if (_logAction != null && !string.IsNullOrWhiteSpace(message))
            {
                _logAction(message);
            }
        }
    }
}
