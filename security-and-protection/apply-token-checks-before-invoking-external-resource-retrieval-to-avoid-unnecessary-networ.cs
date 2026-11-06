using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static async Task Main()
    {
        // Retrieve token (in real scenarios this could come from a secure store)
        string token = GetToken();

        // Validate token before any network operation
        if (!IsTokenValid(token))
        {
            Console.WriteLine("Invalid or missing token. Skipping external resource retrieval.");
            return;
        }

        // URL of the external document (placeholder)
        string url = "https://example.com/sample.docx";

        // Download the document using the valid token
        byte[] documentBytes = await DownloadDocumentAsync(url, token);

        // Load the document from the downloaded bytes
        using (MemoryStream stream = new MemoryStream(documentBytes))
        {
            Document doc = new Document(stream);

            // Apply read‑only protection with a password
            doc.Protect(ProtectionType.ReadOnly, "protectPassword");

            // Save the protected document locally
            string outputPath = "ProtectedDocument.docx";
            doc.Save(outputPath);

            // Verify that the file was created
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The protected document was not saved correctly.");

            Console.WriteLine($"Document saved and protected at '{outputPath}'.");
        }
    }

    // Simulate token retrieval; replace with real implementation as needed
    private static string GetToken()
    {
        // Return empty string to demonstrate token check failure.
        // Change to a non‑empty value to allow the download to proceed.
        return "";
    }

    // Simple token validation logic
    private static bool IsTokenValid(string token)
    {
        return !string.IsNullOrWhiteSpace(token);
    }

    // Perform the HTTP GET request with the token as a Bearer token
    private static async Task<byte[]> DownloadDocumentAsync(string url, string token)
    {
        using (HttpClient client = new HttpClient())
        {
            client.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            HttpResponseMessage response = await client.GetAsync(url);
            response.EnsureSuccessStatusCode();

            return await response.Content.ReadAsByteArrayAsync();
        }
    }
}
