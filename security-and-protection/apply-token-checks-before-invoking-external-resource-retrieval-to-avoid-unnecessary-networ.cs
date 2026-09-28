using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Simple token validation: non‑empty and starts with "valid"
    private static bool IsTokenValid(string token)
    {
        return !string.IsNullOrWhiteSpace(token) && token.StartsWith("valid", StringComparison.OrdinalIgnoreCase);
    }

    // Entry point
    public static async Task Main(string[] args)
    {
        // Retrieve token from environment (or use a placeholder)
        string token = Environment.GetEnvironmentVariable("API_TOKEN") ?? "invalid-token";

        // Decide which workflow to execute based on token validity
        if (IsTokenValid(token))
        {
            // Token is valid – proceed to retrieve external content
            string externalUrl = "https://www.w3.org/TR/PNG/iso_8859-1.txt";

            string content;
            using (HttpClient client = new HttpClient())
            {
                HttpResponseMessage response = await client.GetAsync(externalUrl);
                response.EnsureSuccessStatusCode();
                content = await response.Content.ReadAsStringAsync();
            }

            // Create a new document and insert the retrieved content
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Content retrieved from external resource:");
            builder.Writeln(content);

            // Apply read‑only protection with a password
            doc.Protect(ProtectionType.ReadOnly, "securePassword");

            // Save the protected document
            string outputPath = "ProtectedDocument.docx";
            doc.Save(outputPath, SaveFormat.Docx);

            // Verify that the file was created
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create '{outputPath}'.");
        }
        else
        {
            // Token is invalid – skip external call and create a minimal document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Token validation failed. No external content was retrieved.");

            // Save the unprotected document
            string outputPath = "NoContentDocument.docx";
            doc.Save(outputPath, SaveFormat.Docx);

            // Verify that the file was created
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create '{outputPath}'.");
        }
    }
}
