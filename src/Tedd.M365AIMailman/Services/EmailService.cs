using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

using Tedd.M365AIMailman.Models;

namespace Tedd.M365AIMailman.Services;

internal class EmailService
{
    private readonly ILogger<EmailService> _logger;
    private readonly GraphService _graphService;
    private readonly EmailProcessingSettings _settings;
    private static readonly char[] FolderPathSeparators = new[] { '\\', '/' }; // Accept both separators

    // Category names
    private const string AiReviewedCategoryName = "✓ AI";

    public static string ShortenMessageId(string messageId) =>
        messageId.Length > 8 ? messageId.Substring(0, 8) : messageId;

    public EmailService(ILogger<EmailService> logger, GraphService graphService, IOptions<AppSettings> appSettings)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _graphService = graphService ?? throw new ArgumentNullException(nameof(graphService));

        _settings = appSettings?.Value?.EMailProcessing
            ?? throw new ArgumentNullException(nameof(appSettings), "AppSettings or EMailProcessing section is missing.");

        ArgumentException.ThrowIfNullOrEmpty(_settings.TargetUserId, nameof(_settings.TargetUserId));
        ArgumentException.ThrowIfNullOrEmpty(_settings.SourceFolderName, nameof(_settings.SourceFolderName)); // Path is also required
    }

    /// <summary>
    /// Fetch emails that have NOT yet been reviewed by AI (no "✓ AI" category),
    /// within the configured age window.
    /// </summary>
    public async Task<List<Message>> FetchUnreadEmailsAsync(CancellationToken cancellationToken = default)
    {
        _logger.LogInformation(
            "Attempting to fetch up to {MaxEmails} emails pending AI review for user '{TargetUser}' from path '{SourceFolderPath}'.",
            _settings.MaxEmailsToProcessPerRun, _settings.TargetUserId, _settings.SourceFolderName);

        string? sourceFolderId = null;
        try
        {
            var graphClient = await _graphService.GetAuthenticatedGraphClientAsync();

            // --- Resolve Folder Path to ID ---
            var sourceFolder = await FindFolderByPathAsync(
                graphClient,
                _settings.TargetUserId,
                _settings.SourceFolderName,
                cancellationToken);

            if (sourceFolder?.Id == null)
            {
                _logger.LogError(
                    "Could not find or access the source folder path '{SourceFolderPath}' for user '{TargetUser}'. Please check configuration and permissions.",
                    _settings.SourceFolderName, _settings.TargetUserId);
                return new List<Message>(); // Folder not found or error occurred during lookup
            }

            sourceFolderId = sourceFolder.Id;

            _logger.LogInformation(
                "Successfully resolved source folder path '{SourceFolderPath}' to ID '{FolderId}'. Fetching messages...",
                _settings.SourceFolderName, sourceFolderId);

            // --- Build the Filter ---
            // Only messages NOT yet tagged as AI reviewed
            var filter = $"not(categories/any(c:c eq '{AiReviewedCategoryName}'))";

            // Max age filter
            var cutoffDateMax = DateTimeOffset.UtcNow - _settings.MaxEmailAge;
            var formattedCutoffDateMax = cutoffDateMax.ToString("o"); // "o" is the round-trip format specifier
            filter += $" and receivedDateTime ge {formattedCutoffDateMax}";

            // Min age filter
            var cutoffDateMin = DateTimeOffset.UtcNow - _settings.MinEmailAge;
            var formattedCutoffDateMin = cutoffDateMin.ToString("o");
            filter += $" and receivedDateTime le {formattedCutoffDateMin}";

            _logger.LogInformation(
                "Applying time filter: Fetching emails received on or after {CutoffDateMax}, but before {CutoffDateMin}.",
                formattedCutoffDateMax, formattedCutoffDateMin);

            // Use the resolved Folder ID
            var messages = await graphClient.Users[_settings.TargetUserId]
                .MailFolders[sourceFolderId]
                .Messages
                .GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Filter = filter;
                    requestConfiguration.QueryParameters.Top = _settings.MaxEmailsToProcessPerRun;
                    requestConfiguration.QueryParameters.Select = new[]
                    {
                        "id", "subject", "sender", "from",
                        "bodyPreview", "body", "receivedDateTime", "parentFolderId",
                        "isRead", "categories", "importance", "flag"
                    };
                    requestConfiguration.QueryParameters.Orderby = new[] { "receivedDateTime desc" };
                }, cancellationToken);

            if (messages?.Value != null && messages.Value.Any())
            {
                _logger.LogInformation(
                    "Fetched {Count} emails pending AI review for user '{TargetUser}' from folder ID '{FolderId}' (Path: '{SourceFolderPath}').",
                    messages.Value.Count, _settings.TargetUserId, sourceFolderId, _settings.SourceFolderName);
                return messages.Value;
            }
            else
            {
                _logger.LogInformation(
                    "No emails pending AI review found for user '{TargetUser}' in folder ID '{FolderId}' (Path: '{SourceFolderPath}').",
                    _settings.TargetUserId, sourceFolderId, _settings.SourceFolderName);
                return new List<Message>();
            }
        }
        catch (ODataError odataError)
        {
            _logger.LogError(
                odataError,
                "OData Error fetching emails pending AI review for user '{TargetUser}' from '{SourceFolderPath}' (Resolved ID: {FolderId}). Code: {ErrorCode}, Message: {ErrorMessage}",
                _settings.TargetUserId,
                _settings.SourceFolderName,
                sourceFolderId ?? "N/A",
                odataError.Error?.Code,
                odataError.Error?.Message);
            return new List<Message>();
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Generic error fetching emails pending AI review for user '{TargetUser}' from '{SourceFolderPath}' (Resolved ID: {FolderId}).",
                _settings.TargetUserId,
                _settings.SourceFolderName,
                sourceFolderId ?? "N/A");
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

        var shortMessageId = ShortenMessageId(messageId);
        var actionDescription = markAsRead ? "read" : "unread";

        _logger.LogDebug(
            "Attempting to mark message {MessageId} as {Action} for user '{TargetUser}'.",
            shortMessageId, actionDescription, _settings.TargetUserId);

        try
        {
            var graphClient = await _graphService.GetAuthenticatedGraphClientAsync();
            var update = new Message { IsRead = markAsRead }; // Set IsRead based on the parameter

            // Optional: Pre-check if the message exists
            try
            {
                var checkMessage = await graphClient.Users[_settings.TargetUserId]
                    .Messages[messageId]
                    .GetAsync(requestConfiguration =>
                    {
                        requestConfiguration.QueryParameters.Select = new[] { "id", "isRead" };
                    }, cancellationToken: cancellationToken);

                _logger.LogInformation(
                    "Pre-PATCH check: Message {MessageId} found (current IsRead: {IsReadState}) for user {UserId}.",
                    shortMessageId, checkMessage?.IsRead, _settings.TargetUserId);

                if (checkMessage?.IsRead == markAsRead)
                {
                    _logger.LogInformation(
                        "Message {MessageId} is already marked as {Action}. Skipping redundant PATCH.",
                        shortMessageId, actionDescription);
                    return true;
                }
            }
            catch (ODataError odataGetError) when (odataGetError.ResponseStatusCode == 404)
            {
                _logger.LogWarning(
                    "Pre-PATCH check: Message {MessageId} not found (404) for user {UserId}. Cannot set read state.",
                    shortMessageId, _settings.TargetUserId);
                return false;
            }
            catch (Exception getEx)
            {
                _logger.LogError(
                    getEx,
                    "Pre-PATCH check: Error attempting to GET message {MessageId} for user {UserId}. Proceeding with PATCH attempt.",
                    shortMessageId, _settings.TargetUserId);
            }

            await graphClient.Users[_settings.TargetUserId]
                .Messages[messageId]
                .PatchAsync(update, cancellationToken: cancellationToken);

            _logger.LogInformation(
                "Successfully marked message {MessageId} as {Action} for user '{TargetUser}'.",
                shortMessageId, actionDescription, _settings.TargetUserId);
            return true;

        }
        catch (ODataError odataError)
        {
            _logger.LogError(
                odataError,
                "OData Error marking message {MessageId} as {Action} for user '{TargetUser}'. Status Code: {StatusCode}, Code: {ErrorCode}, Message: {ErrorMessage}",
                shortMessageId,
                actionDescription,
                _settings.TargetUserId,
                odataError.ResponseStatusCode,
                odataError.Error?.Code,
                odataError.Error?.Message);
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Generic error marking message {MessageId} as {Action} for user '{TargetUser}'.",
                shortMessageId, actionDescription, _settings.TargetUserId);
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
    /// Add one or more categories (labels) to a message. Existing categories are preserved and merged.
    /// </summary>
    public Task<bool> AddCategoriesAsync(
        string messageId,
        CancellationToken cancellationToken = default,
        params string[] categoriesToAdd)
    {
        return AddCategoriesAsync(messageId, (IEnumerable<string>)categoriesToAdd, cancellationToken);
    }

    /// <summary>
    /// Core implementation for adding categories to a message.
    /// </summary>
    public async Task<bool> AddCategoriesAsync(
        string messageId,
        IEnumerable<string> categoriesToAdd,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(messageId, nameof(messageId));
        if (categoriesToAdd == null)
            throw new ArgumentNullException(nameof(categoriesToAdd));

        var categoriesList = categoriesToAdd
            .Where(c => !string.IsNullOrWhiteSpace(c))
            .ToList();

        if (categoriesList.Count == 0)
            return true; // Nothing to add

        var shortMessageId = ShortenMessageId(messageId);

        _logger.LogDebug(
            "Adding categories {Categories} to message {MessageId}.",
            string.Join(", ", categoriesList), shortMessageId);

        try
        {
            var graphClient = await _graphService.GetAuthenticatedGraphClientAsync();

            // Get existing categories
            var message = await graphClient.Users[_settings.TargetUserId]
                .Messages[messageId]
                .GetAsync(rc =>
                {
                    rc.QueryParameters.Select = new[] { "id", "categories" };
                }, cancellationToken);

            var merged = new HashSet<string>(
                message?.Categories ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            foreach (var c in categoriesList)
                merged.Add(c);

            var update = new Message
            {
                Categories = merged.ToList()
            };

            await graphClient.Users[_settings.TargetUserId]
                .Messages[messageId]
                .PatchAsync(update, cancellationToken: cancellationToken);

            _logger.LogInformation(
                "Successfully updated categories for message {MessageId}: {Categories}.",
                shortMessageId, string.Join(", ", merged));

            return true;
        }
        catch (ODataError odataError)
        {
            _logger.LogError(
                odataError,
                "OData Error adding categories to message {MessageId}. Status Code: {StatusCode}, Code: {ErrorCode}, Message: {ErrorMessage}",
                shortMessageId,
                odataError.ResponseStatusCode,
                odataError.Error?.Code,
                odataError.Error?.Message);
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Generic error adding categories to message {MessageId}.",
                shortMessageId);
            return false;
        }
    }

    /// <summary>
    /// Tag message as AI-reviewed ("✓ AI").
    /// </summary>
    public Task<bool> TagAiReviewedAsync(string messageId, CancellationToken token = default)
        => AddCategoriesAsync(messageId, token, AiReviewedCategoryName);

    ///// <summary>
    ///// Tag message as AI-reviewed and "Newsletter".
    ///// </summary>
    //public Task<bool> TagNewsletterAsync(string messageId, CancellationToken token = default)
    //    => AddCategoriesAsync(messageId, token, AiReviewedCategoryName, NewsletterCategoryName);

    /// <summary>
    /// Set the Outlook importance (priority) of a message.
    /// </summary>
    public async Task<bool> SetImportanceAsync(
        string messageId,
        Importance importance,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(messageId, nameof(messageId));

        var shortMessageId = ShortenMessageId(messageId);
        _logger.LogDebug(
            "Setting importance of message {MessageId} to {Importance}.",
            shortMessageId, importance);

        try
        {
            var graphClient = await _graphService.GetAuthenticatedGraphClientAsync();

            var update = new Message
            {
                Importance = importance
            };

            await graphClient.Users[_settings.TargetUserId]
                .Messages[messageId]
                .PatchAsync(update, cancellationToken: cancellationToken);

            _logger.LogInformation(
                "Importance of message {MessageId} set to {Importance}.",
                shortMessageId, importance);

            return true;
        }
        catch (ODataError odataError)
        {
            _logger.LogError(
                odataError,
                "OData Error setting importance for message {MessageId}. Status Code: {StatusCode}, Code: {ErrorCode}, Message: {ErrorMessage}",
                shortMessageId,
                odataError.ResponseStatusCode,
                odataError.Error?.Code,
                odataError.Error?.Message);
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Generic error setting importance for message {MessageId}.",
                shortMessageId);
            return false;
        }
    }

    /// <summary>
    /// Flag a message for follow-up (visual flag in Outlook).
    /// </summary>
    public async Task<bool> FlagForFollowUpAsync(
        string messageId,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(messageId, nameof(messageId));

        var shortMessageId = ShortenMessageId(messageId);
        _logger.LogDebug("Flagging message {MessageId} for follow-up.", shortMessageId);

        try
        {
            var graphClient = await _graphService.GetAuthenticatedGraphClientAsync();

            var update = new Message
            {
                Flag = new FollowupFlag
                {
                    FlagStatus = FollowupFlagStatus.Flagged
                }
            };

            await graphClient.Users[_settings.TargetUserId]
                .Messages[messageId]
                .PatchAsync(update, cancellationToken: cancellationToken);

            _logger.LogInformation(
                "Message {MessageId} flagged for follow-up.",
                shortMessageId);

            return true;
        }
        catch (ODataError odataError)
        {
            _logger.LogError(
                odataError,
                "OData Error flagging message {MessageId}. Status Code: {StatusCode}, Code: {ErrorCode}, Message: {ErrorMessage}",
                shortMessageId,
                odataError.ResponseStatusCode,
                odataError.Error?.Code,
                odataError.Error?.Message);
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Generic error flagging message {MessageId}.",
                shortMessageId);
            return false;
        }
    }

    /// <summary>
    /// Clears the follow-up flag on a message.
    /// </summary>
    public async Task<bool> ClearFollowUpFlagAsync(
        string messageId,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(messageId, nameof(messageId));

        var shortMessageId = ShortenMessageId(messageId);
        _logger.LogDebug("Clearing follow-up flag for message {MessageId}.", shortMessageId);

        try
        {
            var graphClient = await _graphService.GetAuthenticatedGraphClientAsync();

            var update = new Message
            {
                Flag = new FollowupFlag
                {
                    FlagStatus = FollowupFlagStatus.NotFlagged
                }
            };

            await graphClient.Users[_settings.TargetUserId]
                .Messages[messageId]
                .PatchAsync(update, cancellationToken: cancellationToken);

            _logger.LogInformation(
                "Follow-up flag cleared for message {MessageId}.",
                shortMessageId);

            return true;
        }
        catch (ODataError odataError)
        {
            _logger.LogError(
                odataError,
                "OData Error clearing follow-up flag for message {MessageId}. Status Code: {StatusCode}, Code: {ErrorCode}, Message: {ErrorMessage}",
                shortMessageId,
                odataError.ResponseStatusCode,
                odataError.Error?.Code,
                odataError.Error?.Message);
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Generic error clearing follow-up flag for message {MessageId}.",
                shortMessageId);
            return false;
        }
    }

    /// <summary>
    /// Finds a MailFolder object by traversing a path like "Folder/SubFolder".
    /// Returns null if the folder path is not found or an error occurs.
    /// Handles well-known folder names like 'Inbox' directly.
    /// </summary>
    private async Task<MailFolder?> FindFolderByPathAsync(
        GraphServiceClient client,
        string userId,
        string folderPath,
        CancellationToken token)
    {
        ArgumentException.ThrowIfNullOrEmpty(folderPath, nameof(folderPath));

        // Handle well-known folders directly for efficiency and common cases
        if (folderPath.Equals("Inbox", StringComparison.OrdinalIgnoreCase))
        {
            _logger.LogDebug("Resolving well-known folder 'Inbox'.");
            return new MailFolder { Id = "Inbox", DisplayName = "Inbox" };
        }

        _logger.LogDebug(
            "Attempting to resolve folder path '{FolderPath}' for user '{UserId}'.",
            folderPath, userId);

        var pathSegments = folderPath.Split(FolderPathSeparators, StringSplitOptions.RemoveEmptyEntries);
        if (pathSegments.Length == 0)
        {
            _logger.LogWarning(
                "Folder path '{FolderPath}' resulted in zero segments.",
                folderPath);
            return null;
        }

        MailFolder? currentFolder = null;
        string? parentFolderId = null; // Start search from root

        try
        {
            for (int i = 0; i < pathSegments.Length; i++)
            {
                string segment = pathSegments[i];
                _logger.LogDebug(
                    "Searching for segment '{Segment}' under parent ID '{ParentId}'.",
                    segment, parentFolderId ?? "root");

                MailFolderCollectionResponse? results;
                if (parentFolderId == null)
                {
                    // Search at the root level (/mailFolders)
                    results = await client.Users[userId].MailFolders
                        .GetAsync(config =>
                        {
                            config.QueryParameters.Filter = $"displayName eq '{segment}'";
                            config.QueryParameters.Select = new[] { "id", "displayName" };
                            config.QueryParameters.Top = 1;
                        }, token);
                }
                else
                {
                    // Search within the child folders of the parent (/mailFolders/{parentFolderId}/childFolders)
                    results = await client.Users[userId].MailFolders[parentFolderId].ChildFolders
                        .GetAsync(config =>
                        {
                            config.QueryParameters.Filter = $"displayName eq '{segment}'";
                            config.QueryParameters.Select = new[] { "id", "displayName" };
                            config.QueryParameters.Top = 1;
                        }, token);
                }

                currentFolder = results?.Value?.FirstOrDefault();

                if (currentFolder == null)
                {
                    _logger.LogWarning(
                        "Could not find folder segment '{Segment}' in path '{FolderPath}' under parent ID '{ParentId}'.",
                        segment, folderPath, parentFolderId ?? "root");
                    return null; // Segment not found
                }

                parentFolderId = currentFolder.Id;

                if (i == pathSegments.Length - 1)
                {
                    _logger.LogDebug(
                        "Successfully resolved last segment '{Segment}' to folder ID {FolderId}.",
                        segment, parentFolderId);
                    return currentFolder;
                }
            }

            _logger.LogError("Folder path resolution logic error for path '{FolderPath}'.", folderPath);
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Error resolving folder path '{FolderPath}' for user '{UserId}'.",
                folderPath, userId);
            return null;
        }
    }
}
