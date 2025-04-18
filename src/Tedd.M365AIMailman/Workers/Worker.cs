using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

using System;
using System.Threading;
using System.Threading.Tasks;

using Tedd.M365AIMailman.Models; // For AppSettings
using Tedd.M365AIMailman.Services; // For ProcessService

namespace Tedd.M365AIMailman.Workers; // Logical grouping for workers

    internal class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private readonly ProcessService _processService;
        private readonly EmailProcessingSettings _settings;

        public Worker(ILogger<Worker> logger, ProcessService processService, IOptions<AppSettings> appSettings)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _processService = processService ?? throw new ArgumentNullException(nameof(processService));
            _settings = appSettings?.Value?.EMailProcessing ?? throw new ArgumentNullException(nameof(appSettings.Value.EMailProcessing));

            if (_settings.PollingIntervalSeconds <= 0)
            {
                _logger.LogWarning("PollingIntervalSeconds is configured to {Interval}s. Using default 60s.", _settings.PollingIntervalSeconds);
                _settings.PollingIntervalSeconds = 60; // Set a sensible default if config is invalid
            }
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _logger.LogInformation("M365 AI Mailman Worker starting at: {time}", DateTimeOffset.Now);

            while (!stoppingToken.IsCancellationRequested)
            {
                _logger.LogInformation("Worker running processing cycle at: {time}", DateTimeOffset.Now);

                try
                {
                    await _processService.ExecuteProcessingCycleAsync(stoppingToken);
                }
                catch (Exception ex)
                {
                    // Catch exceptions from the processing service itself to prevent worker crash
                    _logger.LogCritical(ex, "Critical error during ProcessService execution within the worker loop.");
                    // Consider strategy here: Stop worker? Wait longer before retry?
                }


                try
                {
                    var delay = TimeSpan.FromSeconds(_settings.PollingIntervalSeconds);
                    _logger.LogInformation("Worker cycle complete. Waiting for {Delay} before next cycle.", delay);
                    await Task.Delay(delay, stoppingToken);
                }
                catch (OperationCanceledException)
                {
                    // Expected when stoppingToken is signaled during Task.Delay
                    _logger.LogInformation("Worker stopping gracefully due to cancellation during delay.");
                    break; // Exit the loop
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error during worker delay.");
                    // Decide whether to continue or break on delay errors
                }
            }

            _logger.LogInformation("M365 AI Mailman Worker stopping at: {time}", DateTimeOffset.Now);
        }


        public override async Task StopAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("M365 AI Mailman Worker stopping...");
            await base.StopAsync(cancellationToken);
            _logger.LogInformation("M365 AI Mailman Worker stopped.");
        }


    }
