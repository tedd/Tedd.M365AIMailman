using System.IO.Hashing;
using System.Text;

namespace Tedd.M365AIMailman.Helpers;

/// <summary>
/// Provides utility methods for manipulating message identifiers.
/// </summary>
public static class MessageIdTransformer
{
    /// <summary>
    /// Shortens a Microsoft Graph API Message ID for improved log readability 
    /// by computing its CRC32 checksum and returning it as an 8-character hexadecimal string.
    /// </summary>
    /// <param name="messageId">The original Base64-encoded message ID from Graph API.</param>
    /// <returns>
    /// An 8-character uppercase hexadecimal string representing the CRC32 checksum of the input.
    /// Returns "00000000" if the input is null or empty.
    /// </returns>
    /// <remarks>
    /// This transformation prioritizes brevity and log correlation over uniqueness. 
    /// Collisions are possible but should be infrequent enough for typical logging scenarios.
    /// The CRC32 algorithm provides better distribution than simple truncation.
    /// Requires .NET 6 or later for System.IO.Hashing.Crc32.
    /// For older frameworks, a third-party CRC32 implementation would be necessary.
    /// </remarks>
    public static string ShortenMessageId(string messageId)
    {
        // Input validation: Handle null or empty strings gracefully for logging.
        if (string.IsNullOrEmpty(messageId))
        {
            // Return a default fixed-length placeholder.
            return "00000000";
        }

        try
        {
            // Convert the input string to a byte array using a consistent encoding.
            // UTF-8 is the de facto standard and appropriate here.
            byte[] inputBytes = Encoding.UTF8.GetBytes(messageId);

            // Compute the CRC32 hash. 
            // System.IO.Hashing.Crc32 provides a modern, efficient implementation.
            uint crcValue = Crc32.HashToUInt32(inputBytes);

            // Format the 32-bit unsigned integer as an 8-character hexadecimal string,
            // padded with leading zeros if necessary, using uppercase letters.
            // Hexadecimal is often preferred in logs for density and alignment.
            return crcValue.ToString("X8");
        }
        catch (EncoderFallbackException ex)
        {
            // Although unlikely with standard Base64 strings, handle potential encoding issues.
            // Log the error appropriately in a real application.
            Console.Error.WriteLine($"Error encoding message ID for CRC calculation: {ex.Message}");
            return "ERR_ENCD"; // Indicate an encoding error
        }
        catch (Exception ex) // Catch-all for other unexpected issues
        {
            // Log the error appropriately.
            Console.Error.WriteLine($"Unexpected error shortening message ID '{messageId}': {ex.Message}");
            return "ERR_UNXP"; // Indicate an unexpected error
        }
    }

    // Example Usage:
    // public static void Main(string[] args)
    // {
    //     string id1 = "AAMkADQ1MWMyYmMwLTU3NDItNDg4Mi04NDNlLTc5OGFjYmRkODcwOABGAAAAAACYhvCgyjdDT4BdrQijsmBXBwDKxsPk736JS69819w4daokAAAAAAEMAADKxsPk736JS69819w4daokAAi6s6G8AAA=";
    //     string id2 = "AAMkADQ1MWMyYmMwLTU3NDItNDg4Mi04NDNlLTc5OGFjYmRkODcwOABGAAAAAACYhvCgyjdDT4BdrQijsmBXBwDKxsPk736JS69819w4daokAAAAAAEMAADKxsPk736JS69819w4daokAAi6s6E-AAA=";
    //     string emptyId = "";
    //     string nullId = null;
    //
    //     Console.WriteLine($"Original ID 1: {id1}");
    //     Console.WriteLine($"Shortened ID 1: {ShortenMessageId(id1)}"); // Example Output: E38C3B5B
    //
    //     Console.WriteLine($"Original ID 2: {id2}");
    //     Console.WriteLine($"Shortened ID 2: {ShortenMessageId(id2)}"); // Example Output: 4FDDAB4F
    //     
    //     Console.WriteLine($"Original ID (Empty): '{emptyId}'");
    //     Console.WriteLine($"Shortened ID (Empty): {ShortenMessageId(emptyId)}"); // Example Output: 00000000
    //
    //     Console.WriteLine($"Original ID (Null): '{nullId}'");
    //     Console.WriteLine($"Shortened ID (Null): {ShortenMessageId(nullId)}"); // Example Output: 00000000
    // }
}