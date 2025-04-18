using System;
using System.Collections.Concurrent;
using System.ComponentModel;
using System.Linq; // Needed for FirstOrDefault()
using System.Net;
using System.Threading.Tasks;

using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Graph; // Base Graph namespace
using Microsoft.Graph.Models; // Contains MailFolder, Message, etc.
using Microsoft.Graph.Models.ODataErrors;

// Required for specific request body structures if needed, but often inferred by the fluent API path
// using Microsoft.Graph.Users.Item.Messages.Item.Move; // Example if explicit type needed
using Microsoft.SemanticKernel;

using Tedd.M365AIMailman.Helpers;
using Tedd.M365AIMailman.Models;
using Tedd.M365AIMailman.Services;


namespace Tedd.M365AIMailman.Plugins;

internal class EmailPlugin
{
    private readonly GraphService _graphService;
    private readonly ILogger<EmailPlugin> _logger;
    private readonly EmailProcessingSettings _settings;
    private static readonly ConcurrentDictionary<string, string> _folderIdCache = new(StringComparer.OrdinalIgnoreCase);
    private static readonly char[] FolderPathSeparators = new[] { '\\', '/' }; // Accept both separators


    public EmailPlugin(GraphService graphService, ILogger<EmailPlugin> logger, IOptions<AppSettings> appSettings)
    {
        _graphService = graphService ?? throw new ArgumentNullException(nameof(graphService));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _settings = appSettings?.Value?.EMailProcessing ?? throw new ArgumentNullException(nameof(appSettings.Value.EMailProcessing), "AppSettings or EMailProcessing section is missing.");


        // Validate TargetUserId is present as it's crucial for application permissions
        ArgumentException.ThrowIfNullOrEmpty(_settings.TargetUserId, nameof(_settings.TargetUserId));


        // Optional: Validate specific target folders if needed by logic (commented out in original)
        //if (!_settings.TargetFolders.ContainsKey("Deleted") || string.IsNullOrEmpty(_settings.TargetFolders["Deleted"]))
        //    _logger.LogWarning("Target folder key 'Deleted' is missing or empty in EmailProcessingSettings.");
        //if (!_settings.TargetFolders.ContainsKey("Spam") || string.IsNullOrEmpty(_settings.TargetFolders["Spam"]))
        //    _logger.LogWarning("Target folder key 'Spam' is missing or empty in EmailProcessingSettings.");
    }


    // --- Kernel Functions (Commented out functions kept for reference) ---


    //[KernelFunction, Description("Moves an email message to the configured Deleted Items folder.")]
    //public async Task<string> MoveToDeletedItemsAsync(
    //    [Description("The unique identifier of the email message to delete.")] string messageId)
    //{
    //    ArgumentException.ThrowIfNullOrEmpty(messageId, nameof(messageId));
    //    string deletedFolderName = _settings.TargetFolders.GetValueOrDefault("Deleted", "Deleted Items"); // Fallback, but should be configured
    //    _logger.LogInformation("Attempting to move message {MessageId} for user {UserId} to Deleted folder ('{FolderName}')", messageId, _settings.TargetUserId, deletedFolderName);
    //    return await MoveEmailToFolderInternalAsync(messageId, deletedFolderName);
    //}


    //[KernelFunction("move_to_junk"), Description("Moves an email message to the configured Junk Email (Spam) folder.")]
    //public async Task<string> MoveToJunkAsync(
    //    [Description("The unique identifier of the email message to mark as junk.")] string messageId)
    //{
    //    ArgumentException.ThrowIfNullOrEmpty(messageId, nameof(messageId));
    //    string junkFolderName = _settings.TargetFolders.GetValueOrDefault("Spam", "Junk Email"); // Fallback
    //    _logger.LogInformation("Attempting to move message {MessageId} for user {UserId} to Junk folder ('{FolderName}')", messageId, _settings.TargetUserId, junkFolderName);
    //    return await MoveEmailToFolderInternalAsync(messageId, junkFolderName);
    //}


    [KernelFunction("move_to_folder"), Description("Moves an email message to a specific, configured folder (e.g., 'AI/Newsletters'). Use EXACT folder name/path provided.")]
    public async Task<string> MoveToFolderAsync(
        [Description("The unique identifier of the email message to move.")] string messageId,
        [Description("The EXACT name/path of the target folder (e.g., 'AI/Newsletters', 'AI/Receipts') as configured and listed in the prompt.")] string folderName)
    {
        ArgumentException.ThrowIfNullOrEmpty(messageId, nameof(messageId));
        ArgumentException.ThrowIfNullOrEmpty(folderName, nameof(folderName));


        // Optional validation (kept from original)
        if (_settings.TargetFolders != null && !_settings.TargetFolders.ContainsValue(folderName))
        {
            _logger.LogWarning("Attempted move for user {UserId} to folder '{FolderName}' which is not explicitly configured in TargetFolders. Proceeding, but check configuration/prompt.", _settings.TargetUserId, folderName);
        }


        _logger.LogInformation("Attempting to move message {MessageId} for user {UserId} to specific folder '{FolderName}'",MessageIdTransformer.ShortenMessageId( messageId), _settings.TargetUserId, folderName);
        return await MoveEmailToFolderInternalAsync(messageId, folderName);
    }


    // --- Helper Methods ---


    private async Task<string> MoveEmailToFolderInternalAsync(string messageId, string destinationFolderName)
    {
        var shortMessageId = MessageIdTransformer.ShortenMessageId(messageId);
        if (string.IsNullOrEmpty(destinationFolderName))
        {
            _logger.LogError("MoveEmailToFolderInternalAsync called with null or empty destinationFolderName for message {MessageId} and user {UserId}.", shortMessageId, _settings.TargetUserId);
            return "Error: Destination folder name cannot be empty.";
        }


        try
        {
            // GetFolderIdAsync now also needs the UserId context
            string? destinationFolderId = await GetFolderIdAsync(destinationFolderName, _settings.TargetUserId);
            if (string.IsNullOrEmpty(destinationFolderId))
            {
                var errorMsg = $"Target folder '{destinationFolderName}' not found or could not be created for user {_settings.TargetUserId}.";
                _logger.LogError(errorMsg + " Message ID: {MessageId}", shortMessageId);
                return $"Error: {errorMsg}";
            }


            var graphClient = await _graphService.GetAuthenticatedGraphClientAsync();


            // Define the request body for the move operation.
            // The SDK typically derives the correct body type from the fluent request path.
            // Explicit type: Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody
            var moveRequestBody = new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody
            {
                DestinationId = destinationFolderId
            };


            // *** CHANGE: Use .Users[userId] instead of .Me ***
            await graphClient.Users[_settings.TargetUserId]
                             .Messages[messageId]
                             .Move
                             .PostAsync(moveRequestBody);


            _logger.LogInformation("Successfully moved message {MessageId} for user {UserId} to folder '{FolderName}' (ID: {FolderId})", shortMessageId, _settings.TargetUserId, destinationFolderName, destinationFolderId);
            return $"Successfully moved message {shortMessageId} to {destinationFolderName}.";
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError odataError)
        {
            _logger.LogError(odataError, "OData Error moving message {MessageId} for user {UserId} to folder '{FolderName}'. Code: {ErrorCode}, Message: {ErrorMessage}",
                shortMessageId, _settings.TargetUserId, destinationFolderName, odataError.Error?.Code, odataError.Error?.Message);
            return $"Error moving message {messageId}: {odataError.Error?.Code} - {odataError.Error?.Message}";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Generic error moving message {MessageId} for user {UserId} to folder '{FolderName}'", shortMessageId, _settings.TargetUserId, destinationFolderName);
            return $"Error moving message {shortMessageId}: {ex.Message}";
        }
    }


    // Enhanced folder lookup: Handles paths and creation, now requires userId
    private async Task<string?> GetFolderIdAsync(string folderPath, string userId)
    {
        string cacheKey = $"{userId}::{folderPath}";
        if (_folderIdCache.TryGetValue(cacheKey, out var cachedId))
        {
            _logger.LogDebug("Using cached folder ID for user '{UserId}' path '{FolderPath}': {FolderId}", userId, folderPath, cachedId);
            return cachedId;
        }

        _logger.LogInformation("Resolving folder path '{FolderPath}' for user '{UserId}'.", folderPath, userId);

        try
        {
            var graphClient = await _graphService.GetAuthenticatedGraphClientAsync();
            string? currentParentFolderId = null;
            var pathSegments = folderPath.Split(FolderPathSeparators, StringSplitOptions.RemoveEmptyEntries);

            if (!pathSegments.Any())
            {
                _logger.LogWarning("Folder path '{FolderPath}' resulted in zero segments for user '{UserId}'.", folderPath, userId);
                return null;
            }

            MailFolder? targetFolder = null;
            var baseFolderCollectionRequestBuilder = graphClient.Users[userId].MailFolders;

            for (int i = 0; i < pathSegments.Length; i++)
            {
                var segment = pathSegments[i];
                _logger.LogDebug("Processing segment '{Segment}' for user '{UserId}' under parent ID '{ParentId}'", segment, userId, currentParentFolderId ?? "root");

                MailFolder? foundSegmentFolder = null;
                string filter = $"displayName eq '{Uri.EscapeDataString(segment)}'";
                string[] select = { "id", "displayName" }; // Use array initializer shorthand

                // --- Attempt 1: Find the existing folder segment ---
                try
                {
                    MailFolderCollectionResponse? childFoldersResponse;
                    if (currentParentFolderId == null)
                    {
                        // Use the new RequestConfiguration syntax for GetAsync
                        childFoldersResponse = await baseFolderCollectionRequestBuilder.GetAsync(config =>
                        {
                            config.QueryParameters.Filter = filter;
                            config.QueryParameters.Top = 1;
                            config.QueryParameters.Select = select;
                        });
                    }
                    else
                    {
                        // Use the new RequestConfiguration syntax for GetAsync on ChildFolders
                        childFoldersResponse = await baseFolderCollectionRequestBuilder[currentParentFolderId].ChildFolders.GetAsync(config =>
                        {
                            config.QueryParameters.Filter = filter;
                            config.QueryParameters.Top = 1;
                            config.QueryParameters.Select = select;
                        });
                    }
                    foundSegmentFolder = childFoldersResponse?.Value?.FirstOrDefault();
                }
                catch (ODataError odataFindError)
                {
                    _logger.LogError(odataFindError, "OData Error finding folder segment '{Segment}' for user {UserId} under parent '{ParentId}'. Code: {ErrorCode}, Message: {ErrorMessage}",
                        segment, userId, currentParentFolderId ?? "root", odataFindError.Error?.Code, odataFindError.Error?.Message);
                    return null;
                }
                catch (Exception findEx)
                {
                    _logger.LogError(findEx, "Generic error finding folder segment '{Segment}' for user {UserId} under parent '{ParentId}'.",
                        segment, userId, currentParentFolderId ?? "root");
                    return null;
                }

                // --- Attempt 2: If not found, try to create it ---
                if (foundSegmentFolder == null)
                {
                    _logger.LogInformation("Folder segment '{Segment}' not found for user '{UserId}' under parent ID '{ParentId}'. Attempting to create.", segment, userId, currentParentFolderId ?? "root");
                    var newFolder = new MailFolder { DisplayName = segment };
                    try
                    {
                        targetFolder = currentParentFolderId == null
                            ? await baseFolderCollectionRequestBuilder.PostAsync(newFolder)
                            : await baseFolderCollectionRequestBuilder[currentParentFolderId].ChildFolders.PostAsync(newFolder);

                        if (targetFolder?.Id != null)
                        {
                            _logger.LogInformation("Created folder segment '{Segment}' with ID {FolderId} for user '{UserId}' under parent '{ParentId}'", segment, targetFolder.Id, userId, currentParentFolderId ?? "root");
                            currentParentFolderId = targetFolder.Id;
                        }
                        else
                        {
                            _logger.LogError("Failed to create folder segment '{Segment}' for user '{UserId}' under parent '{ParentId}'. PostAsync returned null or folder without ID.", segment, userId, currentParentFolderId ?? "root");
                            return null;
                        }
                    }
                    catch (ODataError odataCreateError) when (odataCreateError.ResponseStatusCode == (int)HttpStatusCode.Conflict && odataCreateError.Error?.Code == "ErrorFolderExists")
                    {
                        // --- Recovery Logic: Creation failed because it already exists ---
                        _logger.LogWarning(odataCreateError, "Creation of segment '{Segment}' failed because it already exists (Code: {ErrorCode}). Attempting to re-fetch the existing folder.", segment, odataCreateError.Error?.Code);
                        try
                        {
                            MailFolderCollectionResponse? retryResponse;
                            // Re-use filter and select defined above
                            if (currentParentFolderId == null)
                            {
                                retryResponse = await baseFolderCollectionRequestBuilder.GetAsync(config => // Use new syntax
                                {
                                    config.QueryParameters.Filter = filter;
                                    config.QueryParameters.Top = 1;
                                    config.QueryParameters.Select = select;
                                });
                            }
                            else
                            {
                                retryResponse = await baseFolderCollectionRequestBuilder[currentParentFolderId].ChildFolders.GetAsync(config => // Use new syntax
                                {
                                    config.QueryParameters.Filter = filter;
                                    config.QueryParameters.Top = 1;
                                    config.QueryParameters.Select = select;
                                });
                            }

                            var existingFolder = retryResponse?.Value?.FirstOrDefault();
                            if (existingFolder?.Id != null)
                            {
                                _logger.LogInformation("Successfully re-fetched existing folder segment '{Segment}' with ID {FolderId} after creation conflict.", segment, existingFolder.Id);
                                targetFolder = existingFolder;
                                currentParentFolderId = existingFolder.Id;
                            }
                            else
                            {
                                _logger.LogError("Critical inconsistency: Creation failed with ErrorFolderExists for segment '{Segment}', but subsequent fetch failed to find it.", segment);
                                return null;
                            }
                        }
                        catch (ODataError refetchODataError) // Catch ODataError specifically
                        {
                            _logger.LogError(refetchODataError, "OData Error during re-fetch of existing folder segment '{Segment}' after creation conflict. Code: {ErrorCode}, Message: {ErrorMessage}",
                                segment, refetchODataError.Error?.Code, refetchODataError.Error?.Message);
                            return null;
                        }
                        catch (Exception refetchEx)
                        {
                            _logger.LogError(refetchEx, "Generic Error during re-fetch of existing folder segment '{Segment}' after creation conflict.", segment);
                            return null;
                        }
                    }
                    catch (ODataError odataCreateError) // Catch other OData errors during creation
                    {
                        _logger.LogError(odataCreateError, "OData Error creating folder segment '{Segment}' for user {UserId} under parent '{ParentId}'. Code: {ErrorCode}, Message: {ErrorMessage}",
                            segment, userId, currentParentFolderId ?? "root", odataCreateError.Error?.Code, odataCreateError.Error?.Message);
                        return null;
                    }
                    catch (Exception createEx) // Catch generic errors during creation
                    {
                        _logger.LogError(createEx, "Generic error creating folder segment '{Segment}' for user {UserId}' under parent '{ParentId}'.", segment, userId, currentParentFolderId ?? "root");
                        return null;
                    }
                }
                else // Folder segment was found initially
                {
                    _logger.LogDebug("Found existing folder segment '{Segment}' for user '{UserId}' with ID {FolderId}", segment, userId, foundSegmentFolder.Id);
                    targetFolder = foundSegmentFolder;
                    currentParentFolderId = foundSegmentFolder.Id;
                }

                if (i == pathSegments.Length - 1)
                {
                    break;
                }
            } // End for loop

            if (targetFolder?.Id != null)
            {
                _logger.LogInformation("Successfully resolved path '{FolderPath}' for user '{UserId}' to Folder ID {FolderId}", folderPath, userId, targetFolder.Id);
                _folderIdCache.TryAdd(cacheKey, targetFolder.Id);
                return targetFolder.Id;
            }
            else
            {
                _logger.LogError("Could not resolve or create the full folder path '{FolderPath}' for user '{UserId}'. Final targetFolder reference was null or lacked an ID.", folderPath, userId);
                return null;
            }
        }
        catch (ODataError odataError) // Catch errors during initial setup/client acquisition
        {
            _logger.LogError(odataError, "OData Error during folder resolution setup for path '{FolderPath}' for user {UserId}. Code: {ErrorCode}, Message: {ErrorMessage}",
                folderPath, userId, odataError.Error?.Code, odataError.Error?.Message);
            return null;
        }
        catch (Exception ex) // Catch generic errors during setup
        {
            _logger.LogError(ex, "Generic error during folder resolution setup for path '{FolderPath}' for user {UserId}'", folderPath, userId);
            return null;
        }
    }
}