using AssignmentScheduler;
using log4net.Config;
using log4net.Repository;
using log4net;
using System.ComponentModel;

namespace AssignmentScheduler
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var host = CreateHostBuilder(args).Build();

            // Configure log4net
            var configuration = host.Services.GetRequiredService<IConfiguration>();
            ConfigureLog4Net(configuration);
            s

            host.Run();
        }

        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)
                .ConfigureAppConfiguration((hostingContext, config) =>
                {
                    config.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
                })
                .ConfigureServices((hostContext, services) =>
                {
                   
                    services.AddSingleton<s>
                    services.AddHostedService<Worker>();
                });


        private static void ConfigureLog4Net(IConfiguration configuration)
        {
            var log4NetConfig = configuration.GetSection("Logging:Log4NetCore");
            var log4NetConfigFileName = log4NetConfig["Log4NetConfigFileName"];

            ILoggerRepository logRepository = LogManager.GetRepository(System.Reflection.Assembly.GetEntryAssembly());
            XmlConfigurator.ConfigureAndWatch(logRepository, new FileInfo(log4NetConfigFileName));
        }
    }
}

