using log4net.Config;
using log4net.Repository;
using log4net;
using AssignmentScheduler.Interfaces;
using AssignmentScheduler.Repositories;
using AssignmentScheduler.Services;
using Microsoft.Extensions.Options;
using AssignmentScheduler.Configuration;
using MongoDB.Driver;

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
               // Register MongoDB settings
               services.Configure<MongoSettings>(
                   hostContext.Configuration.GetSection(nameof(MongoSettings)));

               services.AddTransient<MongoSettings>(sp =>
                   sp.GetRequiredService<IOptions<MongoSettings>>().Value);

               // Create and register IMongoClient and IMongoDatabase
               services.AddSingleton<IMongoClient, MongoClient>(sp =>
               {
                   var settings = sp.GetRequiredService<IOptions<MongoSettings>>().Value;
                   return new MongoClient(settings.ConnectionString);
               });

               services.AddScoped(sp =>
               {
                   var client = sp.GetRequiredService<IMongoClient>();
                   var settings = sp.GetRequiredService<IOptions<MongoSettings>>().Value;
                   return client.GetDatabase(settings.DatabaseName);
               });

               #region Repositories
               services.AddTransient<IAssignmentRepository, AssignmentRepository>();
               services.AddTransient<IMenRepository, MenRepository>();
               services.AddTransient<IMonthRepository, MonthRepository>();
               services.AddTransient<IProfileRepository, ProfileRepository>();
               services.AddTransient<IProfileAssignmentRepository, ProfileAssignmentRepository>();
               #endregion

               #region Services
               services.AddTransient<IAssignmentExcelGenerator, AssignmentExcelGenerator>();
               services.AddTransient<IEmailService, EmailService>();
               services.AddSingleton<IAwsS3>(provider =>
               {
                   var configuration = provider.GetRequiredService<IConfiguration>();
                   var bucketName = configuration["AWS:BucketName"];
                   var accessKey = configuration["AWS:AccessKey"];
                   var secretKey = configuration["AWS:SecretKey"];
                   var region = configuration["AWS:Region"];

                   return new AwsS3Service(bucketName, accessKey, secretKey, region);
               });
               #endregion

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

