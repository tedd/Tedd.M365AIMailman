using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Graph.Models;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.Connectors.OpenAI;
using Tedd.M365AIMailman.Helpers;
using Tedd.M365AIMailman.Models;
using Tedd.M365AIMailman.Plugins;

namespace Tedd.M365AIMailman.Services;

internal class AIService
{
    private readonly ILogger<AIService> _logger;
    private readonly Kernel _kernel;
    private readonly EmailPlugin _emailPlugin; // Inject the plugin
    private readonly EmailProcessingSettings _emailProcessingSettings;
    private const string EmailPromptTemplate = @"
Analyze the following email details:
From: {{$sender}}
Subject: {{$subject}}
Body Preview: {{$bodyPreview}}
=== BODY START ===
{{$body}}
=== BODY END ===

Based on the content, decide the most appropriate action:
1. Move emails to a Specific Folder: If the email belongs in one of the following categories/folders: {DYNAMIC_FOLDER_LIST}. Use the 'MoveToFolderAsync' function, providing the EXACT folder name from the list above (e.g., '{DYNAMIC_EXAMPLE_FOLDER}' for a newsletter).
3. Emails about refund, pending payments, missing payment or renewals of subscriptions should not be touched.
4. Emails about calendar invite should not be touched unless it is cron-job or similar system generated garbage.
5. Security related emails should not be touched.
6. No Action: If the email seems important, personal, work-related, take no action.

Execute the chosen function. If no action should be taken, respond with 'No action needed'.

To summarize:
* There are emails you should not touch. You need to identify them.
* There are emails you should move to a specific folder. You need to identify them.

Input Email ID: {{$messageId}}
";
    /*
=== BODY START ===
{{$body}}
=== BODY END ===
     */
    
    // Inject Kernel and EmailPlugin directly - they should be registered in Startup
    public AIService(ILogger<AIService> logger, Kernel kernel, EmailPlugin emailPlugin, IOptions<AppSettings> appSettings)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _kernel = kernel ?? throw new ArgumentNullException(nameof(kernel));
        _emailPlugin = emailPlugin ?? throw new ArgumentNullException(nameof(emailPlugin));
        _emailProcessingSettings = appSettings?.Value?.EMailProcessing ?? throw new ArgumentNullException(nameof(appSettings.Value.EMailProcessing));


        // Plugin registration should happen when Kernel is built or here if needed,
        // but preferably during Kernel setup in Startup.cs for clarity.
        // Let's assume it's registered during Kernel building in Startup.
        _logger.LogInformation("AIService initialized with Kernel and EmailPlugin.");
    }

    // Method to process a single email
    public async Task<string> ProcessEmailAsync(Message message, CancellationToken cancellationToken = default)
    {

        if (message is null || string.IsNullOrEmpty(message.Id))
        {
            _logger.LogWarning("ProcessEmailAsync called with null or invalid message.");
            return "Error: Invalid message data.";
        }
        
        _logger.LogInformation("AI Processing Message ID: {MessageId} (short id: {shortId}, Subject: '{Subject}'", message.Id, MessageIdTransformer.ShortenMessageId(message.Id),message.Subject);

        var executionSettings = new OpenAIPromptExecutionSettings
        {
            ToolCallBehavior = ToolCallBehavior.AutoInvokeKernelFunctions
        };

        // Dynamically build folder list (same as before)
        var customFolderNames = _emailProcessingSettings.TargetFolders
            //.Where(kvp => !kvp.Key.Equals("Deleted", StringComparison.OrdinalIgnoreCase) &&
            //              !kvp.Key.Equals("Spam", StringComparison.OrdinalIgnoreCase))
            .Select(kvp => kvp.Value) // Use the actual folder path/name from config
            .ToList();

        string availableFoldersString = string.Join("\t", customFolderNames.Select(f => $"'{f}'"));
        if (string.IsNullOrEmpty(availableFoldersString)) availableFoldersString = "No specific custom folders configured";
        string exampleFolder = _emailProcessingSettings.TargetFolders.GetValueOrDefault("Newsletter", "Mailman/Newsletters"); // Example folder path

        // Prepare the final prompt string using replacement for dynamic C# content
        string emailClassifierPrompt = EmailPromptTemplate // Start with the base template
            .Replace("{DYNAMIC_FOLDER_LIST}", availableFoldersString)
            .Replace("{DYNAMIC_EXAMPLE_FOLDER}", exampleFolder);

        try
        {
            // KernelArguments remain the same, targeting the {{$variable}} placeholders
            var body = HtmlStripper.ExtractText(message.Body?.Content);
            var subject = message.Subject ?? string.Empty;
            body = body.Substring(0, Math.Min(body.Length, 2048));
            subject = subject.Substring(0, Math.Min(subject.Length, 1024));
            var arguments = new KernelArguments(executionSettings)
            {
                { "messageId", message.Id },
                { "sender", message.Sender?.EmailAddress?.Address ?? "Unknown Sender" },
                { "subject", subject },
                { "bodyPreview", message.BodyPreview??string.Empty },
                { "body", body }
            };

            // Add logging for the final prompt to aid diagnostics
            _logger.LogDebug("Final prompt being sent to Kernel:\n{PromptText}", emailClassifierPrompt);

            _logger.LogDebug("Invoking Semantic Kernel for message {MessageId}", MessageIdTransformer.ShortenMessageId(message.Id));
            var result = await _kernel.InvokePromptAsync(emailClassifierPrompt, arguments, cancellationToken: cancellationToken);

            var resultString = result.GetValue<string>() ?? "Kernel returned null or empty result.";
            _logger.LogInformation("Kernel processing result for {MessageId}: {Result}", MessageIdTransformer.ShortenMessageId(message.Id), resultString);

            return resultString;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error invoking Semantic Kernel for message {MessageId}", MessageIdTransformer.ShortenMessageId(message.Id));
            return $"Error: AI processing failed - {ex.Message}";
        }
    }


}