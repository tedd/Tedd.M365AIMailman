using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions.Authentication;

namespace Tedd.M365AIMailman.Helpers;
/// <summary>
/// Implements IAccessTokenProvider using MSAL Confidential Client Flow (Client Credentials).
/// </summary>
internal class MsalClientCredentialProvider : IAccessTokenProvider
{
    private readonly IConfidentialClientApplication _confidentialClientApp;
    private readonly ILogger _logger; // Add logger for diagnostics
    private readonly string[] _scopes = new[] { "https://graph.microsoft.com/.default" };

    public MsalClientCredentialProvider(IConfidentialClientApplication confidentialClientApp, ILogger logger)
    {
        _confidentialClientApp = confidentialClientApp ?? throw new ArgumentNullException(nameof(confidentialClientApp));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <summary>
    /// Gets the authorization token using the confidential client application.
    /// </summary>
    public async Task<string> GetAuthorizationTokenAsync(
        Uri uri,
        Dictionary<string, object>? additionalAuthenticationContext = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            AuthenticationResult authResult = await _confidentialClientApp.AcquireTokenForClient(_scopes)
                .ExecuteAsync(cancellationToken);

            if (authResult != null && !string.IsNullOrEmpty(authResult.AccessToken))
            {
                _logger.LogDebug("MSAL Access Token acquired successfully via MsalClientCredentialProvider.");
                return authResult.AccessToken;
            }
            else
            {
                _logger.LogError("MsalClientCredentialProvider: Failed to acquire token, AuthenticationResult was null or token was empty.");
                // Throwing an exception might be better than returning null/empty
                throw new InvalidOperationException("Failed to acquire application token for Graph API via MSAL.");
            }
        }
        // Catch specific MSAL exceptions if needed for detailed logging, like before
        catch (MsalServiceException ex) when (ex.Message.Contains("AADSTS700016"))
        {
            _logger.LogError(ex, "MsalClientCredentialProvider: Error acquiring token for client (AADSTS700016). Verify Client ID/Tenant ID.");
            throw;
        }
        catch (MsalServiceException ex) when (ex.Message.Contains("AADSTS7000215"))
        {
            _logger.LogError(ex, "MsalClientCredentialProvider: Error acquiring token for client (AADSTS7000215). Verify Client Secret.");
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "MsalClientCredentialProvider: Unexpected error acquiring MSAL token.");
            throw;
        }
    }

    /// <summary>
    /// Gets the allowed hosts validator (usually defaults to graph domains).
    /// </summary>
    public AllowedHostsValidator AllowedHostsValidator => new AllowedHostsValidator(new[] { "graph.microsoft.com", "graph.microsoft.us", "dod-graph.microsoft.us", "graph.microsoft.de", "microsoftgraph.chinacloudapi.cn" });

}
