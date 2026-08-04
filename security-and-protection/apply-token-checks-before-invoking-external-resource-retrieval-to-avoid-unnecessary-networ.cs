using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Entry point of the console application.
    public static async Task Main()
    {
        // Sample token that would normally be obtained from a secure source.
        const string apiToken = "valid-token-123";

        // Validate the token before any network operation.
        if (!IsTokenValid(apiToken))
        {
            Console.WriteLine("Invalid API token. Skipping external resource retrieval.");
            return;
        }

        // URL of a sample Word document. In a real scenario this would be a protected endpoint.
        const string documentUrl = "https://github.com/aspose-words/Aspose.Words-for-.NET/raw/master/Examples/Data/Document.docx";

        // Download the document only after the token has been validated.
        byte[] documentBytes = await DownloadDocumentAsync(documentUrl, apiToken);
        if (documentBytes == null || documentBytes.Length == 0)
        {
            Console.WriteLine("Failed to download the document.");
            return;
        }

        // Load the document from the downloaded byte array using a MemoryStream.
        using (var stream = new MemoryStream(documentBytes))
        {
            Document doc = new Document(stream);

            // Apply read‑only protection with a password.
            const string docPassword = "DocPassword";
            doc.Protect(ProtectionType.ReadOnly, docPassword);

            // Prepare output folder.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // Save the protected document.
            string outputPath = Path.Combine(artifactsDir, "ProtectedDocument.docx");
            doc.Save(outputPath, SaveFormat.Docx);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The protected document was not saved correctly.");

            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
    }

    // Simple token validation logic. Replace with real validation as needed.
    private static bool IsTokenValid(string token)
    {
        // Example rule: token must start with "valid-" and be at least 10 characters long.
        return !string.IsNullOrEmpty(token) && token.StartsWith("valid-") && token.Length >= 10;
    }

    // Downloads a document using HttpClient, passing the token as a bearer token.
    private static async Task<byte[]> DownloadDocumentAsync(string url, string token)
    {
        using (var httpClient = new HttpClient())
        {
            // Add the token to the Authorization header.
            httpClient.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            HttpResponseMessage response = await httpClient.GetAsync(url);
            if (!response.IsSuccessStatusCode)
                return null;

            return await response.Content.ReadAsByteArrayAsync();
        }
    }
}
