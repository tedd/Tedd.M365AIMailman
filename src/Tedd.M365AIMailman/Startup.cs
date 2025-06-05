using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging; // Required for AddLogging
using Microsoft.Extensions.Options;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.Connectors.OpenAI; // Choose appropriate connector
// using Microsoft.SemanticKernel.Connectors.AzureOpenAI;
using Tedd.M365AIMailman.Models;
using Tedd.M365AIMailman.Plugins; // Namespace for EmailPlugin
using Tedd.M365AIMailman.Services;
using Tedd.M365AIMailman.Workers; // Namespace for Worker

namespace Tedd.M365AIMailman;
public static class Startup
{
    public class StartupInt { }
    public static void Initialize(IServiceCollection services, IConfiguration configuration) // Renamed param for clarity
    {
        // --- Configure Options ---
        // Binds configuration sections to strongly typed objects
        services.Configure<AppSettings>(configuration);
        // Makes IOptions<AppSettings> available via DI
        services.AddOptions();


        // --- Register Core Services ---
        // Use Singleton for GraphService if the GraphServiceClient it holds can be reused safely
        // (MSAL handles token refresh internally for Confidential Client)
        //services.AddSingleton<GraphService>();

        // Transient might be suitable for these as they likely hold little state per operation
        services.AddTransient<EmailService>();
        services.AddTransient<AIService>();
        services.AddTransient<ProcessService>();
        services.AddTransient<GraphService>();
        services.AddTransient<EmailPlugin>(); // Register the plugin

        // --- Semantic Kernel Configuration ---
        services.AddSingleton(sp => // Use factory for complex setup
        {
            // ---> Step 1: Get the LoggerFactory configured by the host (which uses Serilog) <---
            var hostLoggerFactory = sp.GetRequiredService<ILoggerFactory>();

            var skSettings = configuration.GetSection("SemanticKernel").Get<SemanticKernelSettings>();

            if (skSettings == null)
            {
                var logger = sp.GetRequiredService<ILogger<StartupInt>>();
                logger.LogCritical("SemanticKernel configuration section is missing or invalid. Application cannot start correctly.");
                throw new InvalidOperationException("SemanticKernel configuration section is missing or invalid.");
            }

            ArgumentException.ThrowIfNullOrEmpty(skSettings.DeploymentOrModelId, "SemanticKernel:DeploymentOrModelId");
            ArgumentException.ThrowIfNullOrEmpty(skSettings.ApiKey, "SemanticKernel:ApiKey");

            // ---> Step 2: Create the Kernel Builder (parameterless) <---
            var kernelBuilder = Kernel.CreateBuilder();

            // ---> Step 3: Add the host's LoggerFactory to the Kernel's internal services <---
            kernelBuilder.Services.AddSingleton(hostLoggerFactory);

            // Add AI Service (Choose based on config)
            if (skSettings.ServiceType.Equals("AzureOpenAI", StringComparison.OrdinalIgnoreCase))
            {
                ArgumentException.ThrowIfNullOrEmpty(skSettings.Endpoint, "SemanticKernel:Endpoint");
                kernelBuilder.AddAzureOpenAIChatCompletion(skSettings.DeploymentOrModelId, skSettings.Endpoint, skSettings.ApiKey);
            }
            else // Default or "OpenAI"
            {
                kernelBuilder.AddOpenAIChatCompletion(skSettings.DeploymentOrModelId, skSettings.ApiKey, skSettings.OrgId); // OrgId is optional
            }

            // Build the kernel - it will now use the hostLoggerFactory for its internal logging needs
            var kernel = kernelBuilder
                .Build();

            // Import the plugin *after* the kernel is built
            var graphSvc = sp.GetRequiredService<GraphService>();
            var pluginLogger = sp.GetRequiredService<ILogger<EmailPlugin>>();
            var appSettings = sp.GetRequiredService<IOptions<AppSettings>>();
            var emailPluginInstance = new EmailPlugin(graphSvc, pluginLogger, appSettings);
            kernel.ImportPluginFromObject(emailPluginInstance, nameof(EmailPlugin));


            var startupLogger = sp.GetRequiredService<ILogger<StartupInt>>();
            startupLogger.LogInformation("Semantic Kernel singleton created and EmailPlugin imported.");

            return kernel;
        });


        // --- Register Hosted Service ---
        services.AddHostedService<Worker>();
    }
}
