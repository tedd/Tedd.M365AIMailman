namespace Tedd.M365AIMailman.Models;
internal class AppSettings
{
    public EmailProcessingSettings EMailProcessing { get; set; } = new();
    public AzureAdSettings AzureAd { get; set; } = new();
    public GraphSettings Graph { get; set; } = new();
    public SemanticKernelSettings SemanticKernel { get; set; } = new();
}