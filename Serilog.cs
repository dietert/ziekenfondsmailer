using Serilog;

namespace ZiekenFondsMailer
{
    static class Serilog
    {
        static readonly ILogger Logger = new LoggerConfiguration()
            .MinimumLevel.Information()
            .WriteTo.Console()
            .CreateLogger();

        public static void Information(string message)
        {
            Logger.Information(message);
        }

        public static void Debug(string message)
        {
            Logger.Debug(message);
        }

        public static void Error(string message)
        {
            Logger.Error(message);
        }

        public static void Warning(string message)
        {
            Logger.Warning(message);
        }

        public static void Fatal(string message)
        {
            Logger.Fatal(message);
        }
    }
}