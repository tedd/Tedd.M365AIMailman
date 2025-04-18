namespace Tedd.M365AIMailman.Models;

internal class AzureAdSettings
{
    public string ClientId { get; set; } = string.Empty;
    public string ClientSecret { get; set; } = string.Empty;
    public string TenantId { get; set; } = string.Empty;
//    public string[] Scopes { get; set; } = new[] { "https://graph.microsoft.com/.default" };
    public string Instance { get; set; } = "https://login.microsoftonline.com/";
}