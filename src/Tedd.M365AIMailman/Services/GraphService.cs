using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
// Required Kiota namespaces for the new auth pattern
using Microsoft.Kiota.Abstractions.Authentication;
using Microsoft.Kiota.Authentication.Azure; // Contains BaseBearerTokenAuthenticationProvider
using Microsoft.Kiota.Http.HttpClientLibrary; // Often needed for GraphServiceClient constructor

using Microsoft.Identity.Client;
using System;
using System.Net.Http.Headers; // Still potentially useful but not directly for auth provider setup
using System.Threading.Tasks;
using Tedd.M365AIMailman.Helpers;
using Tedd.M365AIMailman.Models;


namespace Tedd.M365AIMailman.Services;

internal class GraphService
{
    private readonly ILogger<GraphService> _logger;
    private readonly GraphSettings _graphSettings;
    private readonly AzureAdSettings _azureAdSettings;
    private GraphServiceClient? _graphClient; // Cache the client
    private IConfidentialClientApplication? _confidentialClientApp;

    // Inject ILoggerFactory to pass to MsalClientCredentialProvider if needed
    public GraphService(ILogger<GraphService> logger, ILoggerFactory loggerFactory, IOptions<AppSettings> appSettings)
    {
        // Pass the specific logger for MsalClientCredentialProvider
        // or make MsalClientCredentialProvider require ILogger<MsalClientCredentialProvider>
        // and resolve it via DI if registered separately. For simplicity here, pass loggerFactory.
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _graphSettings = appSettings?.Value?.Graph ?? throw new ArgumentNullException(nameof(appSettings.Value.Graph));
        _azureAdSettings = appSettings?.Value?.AzureAd ?? throw new ArgumentNullException(nameof(appSettings.Value.AzureAd));


        ValidateSettings();
        InitializeConfidentialClient(loggerFactory); // Pass logger factory
    }

    private void ValidateSettings()
    {
        ArgumentException.ThrowIfNullOrEmpty(_azureAdSettings.TenantId, nameof(_azureAdSettings.TenantId));
        ArgumentException.ThrowIfNullOrEmpty(_azureAdSettings.ClientId, nameof(_azureAdSettings.ClientId));
        ArgumentException.ThrowIfNullOrEmpty(_azureAdSettings.ClientSecret, nameof(_azureAdSettings.ClientSecret));
        ArgumentException.ThrowIfNullOrEmpty(_graphSettings.BaseUrl, nameof(_graphSettings.BaseUrl));
    }

    private void InitializeConfidentialClient(ILoggerFactory loggerFactory) // Accept logger factory
    {
        try
        {
            _confidentialClientApp = ConfidentialClientApplicationBuilder
               .Create(_azureAdSettings.ClientId)
               .WithClientSecret(_azureAdSettings.ClientSecret)
               .WithAuthority(new Uri($"{_azureAdSettings.Instance}{_azureAdSettings.TenantId}"))
               // Add MSAL logging if desired
               // .WithLogging((level, message, pii) => {
               //     loggerFactory.CreateLogger("MSAL").LogInformation(message);
               // }, LogLevel.Information, enablePiiLogging: false, enableDefaultPlatformLogging: true)
               .Build();
            _logger.LogInformation("ConfidentialClientApplication initialized successfully.");
        }
        catch (Exception ex)
        {
            _logger.LogCritical(ex, "Failed to initialize ConfidentialClientApplication. Check Azure AD settings.");
            throw new InvalidOperationException("Failed to initialize MSAL Confidential Client Application.", ex);
        }
    }


    // Method to get the client, now async to align potentially with future auth needs
    public GraphServiceClient GetAuthenticatedGraphClient() // Can potentially be non-async now
    {
        if (_graphClient != null)
        {
            return _graphClient;
        }

        if (_confidentialClientApp == null)
        {
            _logger.LogError("ConfidentialClientApplication is not initialized.");
            throw new InvalidOperationException("Cannot get Graph client because ConfidentialClientApplication failed to initialize.");
        }

        // ---> Use the new IAccessTokenProvider implementation <---
        var msalProvider = new MsalClientCredentialProvider(_confidentialClientApp, _logger); // Pass logger

        // ---> Create the BaseBearerTokenAuthenticationProvider <---
        var authProvider = new BaseBearerTokenAuthenticationProvider(msalProvider);

        // ---> Create GraphServiceClient using the Authentication Provider <---
        // Use a standard HttpClient, GraphServiceClient will manage its lifetime if not provided.
        // Or provide a custom configured HttpClient if needed.
        _graphClient = new GraphServiceClient(authProvider);
        // If you need to specify the BaseUrl, you often do it on the request builders now,
        // or configure the underlying HttpClient. Check GraphServiceClient constructor options if needed.


        _logger.LogInformation("GraphServiceClient initialized successfully using BaseBearerTokenAuthenticationProvider.");
        return _graphClient;
    }

    // Keep GetAuthenticatedGraphClientAsync if needed elsewhere, but direct client creation might be sync now
    public async Task<GraphServiceClient> GetAuthenticatedGraphClientAsync()
    {
        // This async method might just call the synchronous one now,
        // unless there was a specific async need for client creation itself.
        await Task.Yield(); // Simulate async work if needed, or just return sync result.
        return GetAuthenticatedGraphClient();
    }
}
