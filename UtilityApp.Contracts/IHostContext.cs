namespace UtilityApp.Contracts
{
    public interface IHostContext
    {
        string ApplicationRootPath { get; }

        string ConfigDirectoryPath { get; }

        string LogsDirectoryPath { get; }

        void Log(string message);
    }
}
