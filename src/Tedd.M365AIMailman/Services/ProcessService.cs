using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

using Tedd.M365AIMailman.Helpers;
using Tedd.M365AIMailman.Models;

namespace Tedd.M365AIMailman.Services;

internal class ProcessService
{
    private readonly ILogger<ProcessService> _logger;
    private readonly EmailService _emailService;
    private readonly AIService _aiService;
    private readonly EmailProcessingSettings _settings;

    public ProcessService(
        ILogger<ProcessService> logger,
        EmailService emailService,
        AIService aiService,
        IOptions<AppSettings> appSettings)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _emailService = emailService ?? throw new ArgumentNullException(nameof(emailService));
        _aiService = aiService ?? throw new ArgumentNullException(nameof(aiService));
        _settings = appSettings?.Value?.EMailProcessing ?? throw new ArgumentNullException(nameof(appSettings.Value.EMailProcessing));
    }

    public async Task ExecuteProcessingCycleAsync(CancellationToken cancellationToken)
    {
        _logger.LogInformation("Starting email processing cycle at {Timestamp}", DateTimeOffset.Now);
        try
        {
            // 1. Fetch Unread Emails
            var messages = await _emailService.FetchUnreadEmailsAsync(cancellationToken);

            if (!messages.Any())
            {
                _logger.LogInformation("No new messages to process in this cycle.");
                return;
            }

            _logger.LogInformation("Processing {Count} emails...", messages.Count);

            // 2. Process Each Email
            foreach (var message in messages)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    _logger.LogWarning("Cancellation requested during processing cycle.");
                    break;
                }

                if (message == null || string.IsNullOrEmpty(message.Id))
                {
                    _logger.LogWarning("Skipping invalid message data received from fetch.");
                    continue;
                }

                var shortMessageId = MessageIdTransformer.ShortenMessageId(message.Id);

                try
                {
                    // 3. Invoke AI Service for classification and action
                    string aiResult = await _aiService.ProcessEmailAsync(message, cancellationToken);

                    // 4. Tag as processed (conditionally based on AI result)
                    // Tag if AI didn't explicitly report an error AND didn't say "No action needed" (implying user should review it as unread).
                    // Adjust this logic based on desired behavior for "No action needed".
                    bool shouldTag = !aiResult.StartsWith("Error:", StringComparison.OrdinalIgnoreCase);
                    // && !aiResult.Equals("No action needed", StringComparison.OrdinalIgnoreCase); // Uncomment if "No action" emails should remain unread

                    _logger.LogInformation("[{MessageId}] Sender: {sender}, Subject: {subject}", shortMessageId, message.Sender?.EmailAddress?.Address, message.Subject);
                    _logger.LogInformation("[{MessageId}] AI conclusion: {message}", shortMessageId, aiResult);
                    _logger.LogInformation("[{MessageId}] Deciding to tag message as {Action} based on AI result.", shortMessageId, shouldTag ? "processed" : "not processed");

                    if (shouldTag)
                    {
                        await _emailService.TagAiReviewedAsync(message.Id, cancellationToken);
                    }
                    else
                    {
                        _logger.LogWarning("Skipping 'AI done' tag for message {MessageId} due to AI result: {Result}", shortMessageId, aiResult);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Unhandled error processing individual message {MessageId}. Skipping message.", shortMessageId);
                    // Decide if you want to attempt marking as read even on error, or leave unread for retry.
                }
            } // End foreach message
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unhandled error during email processing cycle execution.");
            // Consider backoff or other strategies if the whole cycle fails
        }
        finally
        {
            _logger.LogInformation("Finished email processing cycle at {Timestamp}", DateTimeOffset.Now);
        }
    }
}