namespace Tedd.M365AIMailman.Models;

internal class EmailProcessingSettings
{
    public string TargetUserId { get; set; } = string.Empty; // User ID for the target mailbox
    public string SourceFolderName { get; set; } = "Inbox"; // Default source folder
    public TimeSpan MaxEmailAge { get; set; } = new TimeSpan(days: 7,hours:0,minutes:0,seconds:0);
    public TimeSpan MinEmailAge { get; set; } = new TimeSpan(0, 0, 5, 0);
    public Dictionary<string, string> TargetFolders { get; set; } = new()
    {
    };

    public int PollingIntervalSeconds { get; set; } = 300;
    public int MaxEmailsToProcessPerRun { get; set; } = 20;
}