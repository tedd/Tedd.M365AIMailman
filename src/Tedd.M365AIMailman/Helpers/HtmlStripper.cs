using HtmlAgilityPack;

namespace Tedd.M365AIMailman.Helpers;

/// <summary>
/// Provides functionality to extract plain text content from HTML markup.
/// </summary>
public static class HtmlStripper
{
    /// <summary>
    /// Extracts plain text from an HTML string using HtmlAgilityPack.
    /// </summary>
    /// <param name="htmlContent">The HTML string input.</param>
    /// <returns>The extracted plain text, or an empty string if the input is null or whitespace.</returns>
    /// <remarks>
    /// This method leverages HtmlAgilityPack to parse the HTML structure
    /// and retrieve the inner text content, effectively stripping all tags.
    /// It handles null or empty input gracefully.
    /// </remarks>
    public static string ExtractText(string? htmlContent)
    {
        if (string.IsNullOrWhiteSpace(htmlContent))
        {
            // Return empty string for null, empty, or whitespace input
            // to avoid unnecessary processing or potential exceptions.
            return string.Empty;
        }

        HtmlDocument htmlDoc = new HtmlDocument();
        htmlDoc.LoadHtml(htmlContent);

        // The DocumentNode represents the root of the HTML document.
        // Accessing its InnerText property recursively aggregates the text
        // content of all descendant nodes, effectively stripping the tags.
        string extractedText = htmlDoc.DocumentNode.InnerText;

        // Optional: Further refine the text, e.g., decode HTML entities
        // that might remain if not handled by InnerText (though HAP usually does).
        // Consider trimming whitespace or normalizing line breaks if needed.
        // Example: return System.Net.WebUtility.HtmlDecode(extractedText).Trim();

        return extractedText;
    }

   
}