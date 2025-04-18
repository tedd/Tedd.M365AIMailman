using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Serilog;

namespace Tedd.M365AIMailman.ConsoleUI
{
    internal class Program
    {
        static async Task<int> Main(string[] args)
        {
            // --- Serilog Configuration ---
            // Configure Serilog logger first to capture logs during host build process.
            // Read base configuration from appsettings.json, override with environment-specific
            // files (e.g., appsettings.Development.json) and environment variables.
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .AddJsonFile($"appsettings.{Environment.GetEnvironmentVariable("DOTNET_ENVIRONMENT") ?? "Production"}.json", optional: true)
                // Add development explicitly for bootstrap if needed, also optional
                .AddJsonFile("appsettings.Development.json", optional: true)
                .AddEnvironmentVariables()
                .AddCommandLine(args) // Allow command line args to override config
                .AddUserSecrets<Program>()
                .Build();

            // Configure Serilog using the application configuration.
            // This setup reads sinks, enrichment, minimum levels etc., from the "Serilog" section.
            Log.Logger = new LoggerConfiguration()
                .ReadFrom.Configuration(configuration) // Read config from IConfiguration
                .Enrich.FromLogContext() // Enrich logs with context properties
                .CreateBootstrapLogger(); // Use CreateBootstrapLogger for early logging
            try
            {
                // Log application start using Serilog's static logger.
                Log.Information("Initializing Host with...");

            // Create and configure the host builder.
            // Host.CreateDefaultBuilder() provides default configuration for logging,
            // configuration sources (appsettings.json, environment variables, command line), etc.
            var builder = Host.CreateDefaultBuilder(args);

            // --- Customize Host Configuration ---
            builder.ConfigureAppConfiguration((hostContext, configBuilder) =>
            {
                // CreateDefaultBuilder already adds:
                // 1. appsettings.json (required by default)
                // 2. appsettings.{Environment}.json (optional, based on DOTNET_ENVIRONMENT)
                // 3. User Secrets (if Environment is Development)
                // 4. Environment Variables
                // 5. Command Line args

                // Explicitly add appsettings.Development.json AFTER the defaults.
                // If DOTNET_ENVIRONMENT is 'Development', it might technically be added twice,
                // but the last source added wins for duplicate keys, so it's generally safe.
                // Crucially, setting optional: true means it won't crash if the file
                // doesn't exist (e.g., in Production).
                configBuilder.AddJsonFile("appsettings.Development.json",
                    optional: true,  // Don't throw an error if file is missing
                    reloadOnChange: true); // Reload if file changes during runtime
            });

                // --- Serilog Integration with Host ---
                builder.UseSerilog((hostContext, services, loggerConfiguration) => loggerConfiguration
                    .ReadFrom.Configuration(hostContext.Configuration) // Read config again for host context
                    .Enrich.FromLogContext()
                // Add any additional code-based configuration if needed
                // Example: .MinimumLevel.Override("Microsoft", LogEventLevel.Warning)
            );

            // Configure services by calling the static Initialize method from the Startup class.
            // This centralizes the service registration and configuration logic.
            builder.ConfigureServices((hostContext, services) =>
            {
                // Delegate initialization to the static Startup class.
                Startup.Initialize(services, hostContext.Configuration);
            });

            // Build the host instance.
            using var host = builder.Build();

            Log.Information("Host built. Starting application...");

            // Run the host asynchronously.
            await host.RunAsync();

            Log.Information("Application shutting down normally.");
            return 0; // Success exit code
            }
            catch (Exception ex)
            {
                // Log fatal exceptions that occur during startup or execution.
                Log.Fatal(ex, "Host terminated unexpectedly.");
                return 1; // Error exit code
            }
            finally
            {
                // --- Ensure logs are flushed ---
                // Close and flush the Serilog logger before application exit.
                // This is crucial to ensure all buffered logs are written.
                await Log.CloseAndFlushAsync();
            }
        }
    }
}
