using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using SPOFooterSiteConfigListProvisioner.Interfaces;
using System.IO;
using System.Threading.Tasks;
using Serilog;

namespace SPOFooterSiteConfigListProvisioner
{
    class Program
    {
        private static void Main(string[] args)
        {
            CreateHostBuilder(args).Build().Run();
        }

        private static IHostBuilder CreateHostBuilder(string[] args) => Host
            .CreateDefaultBuilder(args)
            .UseConsoleLifetime(o => o.SuppressStatusMessages = true)
            .ConfigureAppConfiguration((hostContext, config) =>
            {
                config
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile($"appsettings.json", optional: false, reloadOnChange: true)
                    .AddEnvironmentVariables()
                    .AddCommandLine(args);

                if (hostContext.HostingEnvironment.IsDevelopment())
                    config.AddUserSecrets<ConfigurationSettings>(); // inject config settings from local secrets.json
            })
            .ConfigureServices((hostContext, services) =>
            {
                services
                    // inject configuration settings
                    .Configure<ConfigurationSettings>(hostContext.Configuration.GetSection(ConfigurationSettings.SectionName))

                    // inject the task scheduler
                    .AddSingleton<TaskScheduler, LimitedConcurrencyLevelTaskScheduler>()
                    
                    // inject the SPO url discovery service
                    .AddSingleton<ISiteUrlDiscoveryService, TenantSiteUrlDiscoveryService>()
                    
                    // inject our list provisioner service
                    .AddSingleton<ISiteCollectionExecuter, SiteConfigListProvisioner>()
                    
                    // finally, inject our generic background task service
                    .AddHostedService<TaskExecutionBackgroundService>();
            })
            .ConfigureLogging((hostContext, logging) =>
            {
                logging
                    .AddConsole()
                    .AddConfiguration(hostContext.Configuration.GetSection("Logging"))
                    .AddFile(hostContext.Configuration.GetSection("Logging"));
            });
    }
}
