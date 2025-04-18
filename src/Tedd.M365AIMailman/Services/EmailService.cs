using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

using Tedd.M365AIMailman.Helpers;
using Tedd.M365AIMailman.Models;

namespace Tedd.M365AIMailman.Services;

internal class EmailService
{
    private readonly ILogger<EmailService> _logger;
    private readonly GraphService _graphService;
    private readonly EmailProcessingSettings _settings;
    private static readonly char[] FolderPathSeparators = new[] { '\\', '/' }; // Accept both separators

    public static string ShortenMessageId (string messageId) => messageId.Length > 8 ? messageId.Substring(0, 8) : messageId;

    public EmailService(ILogger<EmailService> logger, GraphService graphService, IOptions<AppSettings> appSettings)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _graphService = graphService ?? throw new ArgumentNullException(nameof(graphService));

        _settings = appSettings?.Value?.EMailProcessing
            ?? throw new ArgumentNullException(nameof(appSettings), "AppSettings or EMailProcessing section is missing.");

        ArgumentException.ThrowIfNullOrEmpty(_settings.TargetUserId, nameof(_settings.TargetUserId));
        ArgumentException.ThrowIfNullOrEmpty(_settings.SourceFolderName, nameof(_settings.SourceFolderName)); // Path is also required
    }

    public async Task<List<Message>> FetchUnreadEmailsAsync(CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("Attempting to fetch up to {MaxEmails} unread emails for user '{TargetUser}' from path '{SourceFolderPath}'.",
            _settings.MaxEmailsToProcessPerRun, _settings.TargetUserId, _settings.SourceFolderName);

        string? sourceFolderId = null;
        try
        {
            var graphClient = await _graphService.GetAuthenticatedGraphClientAsync();

            // --- Resolve Folder Path to ID ---
            var sourceFolder = await FindFolderByPathAsync(graphClient, _settings.TargetUserId, _settings.SourceFolderName, cancellationToken);

            if (sourceFolder?.Id == null)
            {
                _logger.LogError("Could not find or access the source folder path '{SourceFolderPath}' for user '{TargetUser}'. Please check configuration and permissions.",
                    _settings.SourceFolderName, _settings.TargetUserId);
                return new List<Message>(); // Folder not found or error occurred during lookup
            }
            sourceFolderId = sourceFolder.Id;
            // --- Folder ID resolved ---

            _logger.LogInformation("Successfully resolved source folder path '{SourceFolderPath}' to ID '{FolderId}'. Fetching messages...",
                _settings.SourceFolderName, sourceFolderId);
            // --- Build the Filter ---
            var filter = "isRead eq false"; // Start with the mandatory filter
            
            // Calculate the cutoff date in UTC
            var cutoffDate = DateTimeOffset.UtcNow -_settings.MaxEmailAge;
            // Format for Graph query (ISO 8601 format)
            var formattedCutoffDate1 = cutoffDate.ToString("o"); // "o" is the round-trip format specifier

            // Append the date filter condition
            filter += $" and receivedDateTime ge {formattedCutoffDate1}";


            // And the minimum age filter
            cutoffDate = DateTimeOffset.UtcNow - _settings.MinEmailAge;
            var formattedCutoffDate2 = cutoffDate.ToString("o"); // "o" is the round-trip format specifier
            // Append the date filter condition
            filter += $" and receivedDateTime le {formattedCutoffDate2}";

            
            _logger.LogInformation("Applying time filter: Fetching emails received on or after {CutoffDate}, but before {CutoffDate2}.", formattedCutoffDate1, formattedCutoffDate2);
            // Use the resolved Folder ID
            var messages = await graphClient.Users[_settings.TargetUserId]
                .MailFolders[sourceFolderId] // Use the resolved Folder ID
                .Messages
                .GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Filter = filter;
                    requestConfiguration.QueryParameters.Top = _settings.MaxEmailsToProcessPerRun;
                    requestConfiguration.QueryParameters.Select = new[] {
                        "id", "subject", "sender", "from",
                        "bodyPreview", "body", "receivedDateTime", "parentFolderId", "isRead"
                    };
                    requestConfiguration.QueryParameters.Orderby = new[] { "receivedDateTime desc" };
                    
                    
                }, cancellationToken);

            if (messages?.Value != null && messages.Value.Any())
            {
                _logger.LogInformation("Fetched {Count} unread emails for user '{TargetUser}' from folder ID '{FolderId}' (Path: '{SourceFolderPath}').",
                    messages.Value.Count, _settings.TargetUserId, sourceFolderId, _settings.SourceFolderName);
                return messages.Value;
            }
            else
            {
                _logger.LogInformation("No unread emails found for user '{TargetUser}' in folder ID '{FolderId}' (Path: '{SourceFolderPath}').",
                    _settings.TargetUserId, sourceFolderId, _settings.SourceFolderName);
                return new List<Message>();
            }
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError odataError)
        {
            _logger.LogError(odataError, "OData Error fetching unread emails for user '{TargetUser}' from '{SourceFolderPath}' (Resolved ID: {FolderId}). Code: {ErrorCode}, Message: {ErrorMessage}",
                 _settings.TargetUserId, _settings.SourceFolderName, sourceFolderId ?? "N/A", odataError.Error?.Code, odataError.Error?.Message);
            return new List<Message>();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Generic error fetching unread emails for user '{TargetUser}' from '{SourceFolderPath}' (Resolved ID: {FolderId}).",
                 _settings.TargetUserId, _settings.SourceFolderName, sourceFolderId ?? "N/A");
            return new List<Message>();
        }
    }

    /// <summary>
    /// Sets the read state of a specific email message for the configured target user.
    /// </summary>
    /// <param name="messageId">The unique identifier of the message.</param>
    /// <param name="markAsRead">True to mark the message as read, false to mark it as unread.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <returns>True if the operation was successful, false otherwise.</returns>
    public async Task<bool> SetEmailReadStateAsync(string messageId, bool markAsRead, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(messageId, nameof(messageId));
        // TargetUserId validated in constructor

        var shortMessageId = MessageIdTransformer.ShortenMessageId(messageId);


        string actionDescription = markAsRead ? "read" : "unread";
        _logger.LogDebug("Attempting to mark message {MessageId} as {Action} for user '{TargetUser}'.", messageId, actionDescription, _settings.TargetUserId);

        try
        {
            var graphClient = await _graphService.GetAuthenticatedGraphClientAsync();
            var update = new Message { IsRead = markAsRead }; // Set IsRead based on the parameter

            // Optional: Pre-check if the message exists (helps avoid unnecessary PATCH calls on deleted items)
            try
            {
                // We only need to know it exists, so select minimal fields or just perform the GET.
                // A HEAD request would be ideal but isn't directly supported for specific messages in Graph SDK v5 in this manner easily.
                // GET is acceptable for a pre-check.
                var checkMessage = await graphClient.Users[_settings.TargetUserId]
                    .Messages[messageId]
                    .GetAsync(requestConfiguration => {
                        requestConfiguration.QueryParameters.Select = new[] { "id", "isRead" }; // Minimal select
                    }, cancellationToken: cancellationToken);

                // Log current state if needed for diagnostics
                _logger.LogInformation("Pre-PATCH check: Message {MessageId} found (current IsRead: {IsReadState}) for user {UserId}.", shortMessageId, checkMessage?.IsRead, _settings.TargetUserId);

                // Optional: Add logic to skip PATCH if already in the desired state
                if (checkMessage?.IsRead == markAsRead)
                {
                    _logger.LogInformation("Message {MessageId} is already marked as {Action}. Skipping redundant PATCH.", shortMessageId, actionDescription);
                    return true; // Operation is effectively successful as state is already correct
                }

            }
            catch (ODataError odataGetError) when (odataGetError.ResponseStatusCode == 404)
            {
                _logger.LogWarning("Pre-PATCH check: Message {MessageId} not found (404) for user {UserId}. Cannot set read state.", shortMessageId, _settings.TargetUserId);
                return false; // Message doesn't exist, cannot proceed.
            }
            catch (Exception getEx)
            {
                // Log the error but potentially proceed with the PATCH attempt, as the GET failure might be transient
                _logger.LogError(getEx, "Pre-PATCH check: Error attempting to GET message {MessageId} for user {UserId}. Proceeding with PATCH attempt.", shortMessageId, _settings.TargetUserId);
            }

            // Perform the PATCH operation
            // Using the fluent API for PATCH:
            await graphClient.Users[_settings.TargetUserId]
                .Messages[messageId]
                .PatchAsync(update, cancellationToken: cancellationToken);


            _logger.LogInformation("Successfully marked message {MessageId} as {Action} for user '{TargetUser}'.", shortMessageId, actionDescription, _settings.TargetUserId);
            return true;

        }
        catch (ODataError odataError)
        {
            // Log specific OData errors (e.g., permissions, throttling, not found during PATCH)
            _logger.LogError(odataError, "OData Error marking message {MessageId} as {Action} for user '{TargetUser}'. Status Code: {StatusCode}, Code: {ErrorCode}, Message: {ErrorMessage}",
                shortMessageId, actionDescription, _settings.TargetUserId, odataError.ResponseStatusCode, odataError.Error?.Code, odataError.Error?.Message);
            return false;
        }
        catch (Exception ex)
        {
            // Catch-all for other unexpected errors (network issues, etc.)
            _logger.LogError(ex, "Generic error marking message {MessageId} as {Action} for user '{TargetUser}'.", shortMessageId, actionDescription, _settings.TargetUserId);
            return false;
        }
    }


    /// <summary>
    /// Convenience method to mark an email as read.
    /// </summary>
    public Task<bool> MarkEmailAsReadAsync(string messageId, CancellationToken cancellationToken = default)
    {
        return SetEmailReadStateAsync(messageId, true, cancellationToken);
    }

    /// <summary>
    /// Convenience method to mark an email as unread.
    /// </summary>
    public Task<bool> MarkEmailAsUnreadAsync(string messageId, CancellationToken cancellationToken = default)
    {
        return SetEmailReadStateAsync(messageId, false, cancellationToken);
    }

    /// <summary>
    /// Finds a MailFolder object by traversing a path like "Folder/SubFolder".
    /// Returns null if the folder path is not found or an error occurs.
    /// Handles well-known folder names like 'Inbox' directly.
    /// </summary>
    private async Task<MailFolder?> FindFolderByPathAsync(GraphServiceClient client, string userId, string folderPath, CancellationToken token)
    {
        ArgumentException.ThrowIfNullOrEmpty(folderPath, nameof(folderPath));

        // Handle well-known folders directly for efficiency and common cases
        if (folderPath.Equals("Inbox", StringComparison.OrdinalIgnoreCase))
        {
            _logger.LogDebug("Resolving well-known folder 'Inbox'.");
            // We can return a dummy MailFolder with just the ID if that's all we need downstream,
            // or fetch the actual folder if other properties are needed. For fetching messages,
            // just the well-known name works as the ID.
            return new MailFolder { Id = "Inbox", DisplayName = "Inbox" };
        }
        // Add other well-known folders here if needed (SentItems, Drafts, etc.)
        // else if (folderPath.Equals("Sent Items", StringComparison.OrdinalIgnoreCase)) { return new MailFolder { Id = "SentItems", DisplayName = "Sent Items"}; }

        _logger.LogDebug("Attempting to resolve folder path '{FolderPath}' for user '{UserId}'.", folderPath, userId);

        var pathSegments = folderPath.Split(FolderPathSeparators, StringSplitOptions.RemoveEmptyEntries);
        if (pathSegments.Length == 0)
        {
            _logger.LogWarning("Folder path '{FolderPath}' resulted in zero segments.", folderPath);
            return null;
        }

        MailFolder? currentFolder = null;
        string? parentFolderId = null; // Start search from root

        try
        {
            for (int i = 0; i < pathSegments.Length; i++)
            {
                string segment = pathSegments[i];
                _logger.LogDebug("Searching for segment '{Segment}' under parent ID '{ParentId}'.", segment, parentFolderId ?? "root");

                MailFolderCollectionResponse? results;
                if (parentFolderId == null)
                {
                    // Search at the root level (/mailFolders)
                    results = await client.Users[userId].MailFolders
                        .GetAsync(config => {
                            config.QueryParameters.Filter = $"displayName eq '{segment}'";
                            config.QueryParameters.Select = new[] { "id", "displayName" };
                            config.QueryParameters.Top = 1;
                        }, token);
                }
                else
                {
                    // Search within the child folders of the parent (/mailFolders/{parentFolderId}/childFolders)
                    results = await client.Users[userId].MailFolders[parentFolderId].ChildFolders
                        .GetAsync(config => {
                            config.QueryParameters.Filter = $"displayName eq '{segment}'";
                            config.QueryParameters.Select = new[] { "id", "displayName" };
                            config.QueryParameters.Top = 1;
                        }, token);
                }

                currentFolder = results?.Value?.FirstOrDefault();

                if (currentFolder == null)
                {
                    _logger.LogWarning("Could not find folder segment '{Segment}' in path '{FolderPath}' under parent ID '{ParentId}'.",
                        segment, folderPath, parentFolderId ?? "root");
                    return null; // Segment not found
                }

                parentFolderId = currentFolder.Id; // ID found, use it as parent for the next segment

                if (i == pathSegments.Length - 1)
                {
                    // This is the last segment, we found our target folder
                    _logger.LogDebug("Successfully resolved last segment '{Segment}' to folder ID {FolderId}.", segment, parentFolderId);
                    return currentFolder;
                }
            }

            // Should not be reachable if pathSegments has items, but compiler requires return path
            _logger.LogError("Folder path resolution logic error for path '{FolderPath}'.", folderPath);
            return null;

        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error resolving folder path '{FolderPath}' for user '{UserId}'.", folderPath, userId);
            return null;
        }
    }
}