namespace Tedd.M365AIMailman.Models;

internal class SemanticKernelSettings
{
    
    public string ServiceType { get; set; } = "OpenAI"; // Or "AzureOpenAI"
    public string DeploymentOrModelId { get; set; } = "gpt-o4-mini"; // e.g., gpt-4, gpt-35-turbo
    // --- Azure OpenAI Specific ---
    public string Endpoint { get; set; } =string.Empty; // Required if ServiceType is AzureOpenAI
    // --- OpenAI Specific ---
    public string OrgId { get; set; } = string.Empty; // Optional for OpenAI
    // --- Common ---
    public string ApiKey { get; set; } = string.Empty; // Required
    
}