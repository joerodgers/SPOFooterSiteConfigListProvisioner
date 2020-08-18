using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using SPOFooterSiteConfigListProvisioner.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace SPOFooterSiteConfigListProvisioner
{
    public class TaskExecutionBackgroundService : BackgroundService
    {
        private readonly ILogger<TaskExecutionBackgroundService> _logger;
        private readonly IHostEnvironment _hostEnvironment;
        private readonly IHostApplicationLifetime _applicationLifetime;
        private readonly TaskScheduler _taskScheduler;
        private readonly ISiteCollectionExecuter _siteCollectionExecuter;
        private readonly ISiteUrlDiscoveryService _siteUrlDiscoveryService;

        public TaskExecutionBackgroundService(
            ILogger<TaskExecutionBackgroundService> logger, 
            IHostEnvironment hostEnvironment,
            IHostApplicationLifetime applicationLifetime, 
            TaskScheduler taskScheduler,
            ISiteCollectionExecuter siteCollectionExecuter,
            ISiteUrlDiscoveryService siteUrlDiscoveryService)
        {
            _logger                  = logger                  ?? throw new ArgumentNullException(nameof(logger));
            _hostEnvironment         = hostEnvironment         ?? throw new ArgumentNullException(nameof(hostEnvironment));
            _applicationLifetime     = applicationLifetime     ?? throw new ArgumentNullException(nameof(applicationLifetime));
            _taskScheduler           = taskScheduler           ?? throw new ArgumentNullException(nameof(taskScheduler));
            _siteCollectionExecuter  = siteCollectionExecuter  ?? throw new ArgumentNullException(nameof(siteCollectionExecuter));
            _siteUrlDiscoveryService = siteUrlDiscoveryService ?? throw new ArgumentNullException(nameof(siteUrlDiscoveryService));
        }

        protected override Task ExecuteAsync(CancellationToken stoppingToken)
        {
 
            try
            {
                var counter = 0;
                var tasks = new List<Task>();
                var factory = new TaskFactory(_taskScheduler);
                
                var urls = _siteUrlDiscoveryService.GetSiteCollectionUrls();

                if(_hostEnvironment.IsDevelopment() && urls.Count > 10)
                {
                    urls = urls.Take(10).ToList();
                }

                urls.ForEach(delegate (string url)
                {
                    if( counter > 0 && counter % 5 == 0)
                        _logger.LogInformation($"{DateTime.Now} - Scheduled {counter} tasks");

                    if (counter == urls.Count - 1)
                        _logger.LogInformation($"{DateTime.Now} - Scheduled {urls.Count} tasks");

                    tasks.Add(factory.StartNew(() =>
                    {
                        _siteCollectionExecuter.Execute(url);
                    },
                    stoppingToken));

                    counter++;
                });

                _logger.LogInformation($"{DateTime.Now} - Waiting for the execution of {tasks.Count} tasks.");

                Task.WaitAll(tasks.ToArray());
            }
            catch(Exception ex)
            {
                _logger.LogCritical(ex, $"Error executing tasks.");
            }
            finally 
            {
                _applicationLifetime.StopApplication();
            }

            return Task.CompletedTask;
        }
    }
}
