namespace Tedd.M365AIMailman.Models;

internal class EmailProcessingSettings
{
    public String TargetUserId { get; set; } = String.Empty; // User ID for the target mailbox
    public String SourceFolderName { get; set; } = "Inbox"; // Default source folder
    public TimeSpan MaxEmailAge { get; set; } = new TimeSpan(days: 7, hours: 0, minutes: 0, seconds: 0);
    public TimeSpan MinEmailAge { get; set; } = new TimeSpan(0, 0, 5, 0);
    public Dictionary<String, String> TargetFolders { get; set; } = new()
    {
    };

    public Int32 PollingIntervalSeconds { get; set; } = 300;
    public Int32 MaxEmailsToProcessPerRun { get; set; } = 20;
    public String PromptFile { get; set; } = "Prompt.txt";

    public String SystemPrompt { get; set; } = "You are an AI assistant that classifies emails into specific folders based on their content. Use the provided rules and folder list to determine the best fit for each email.";
}