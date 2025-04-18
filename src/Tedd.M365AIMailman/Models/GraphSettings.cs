namespace Tedd.M365AIMailman.Models;

internal class GraphSettings
{
    public List<string> Scopes { get; set; } = new()
    {
        "Mail.ReadWrite", "User.Read"
    };
    public string BaseUrl { get; set; } = "https://graph.microsoft.com/v1.0";
}