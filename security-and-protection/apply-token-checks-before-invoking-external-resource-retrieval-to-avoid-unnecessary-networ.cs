using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;

public class Program
{
    // Simulated token validation. In real scenarios this could involve JWT validation, expiry checks, etc.
    private static bool IsTokenValid(string token)
    {
        // For demonstration, only the token value "valid-token" is considered valid.
        return token == "valid-token";
    }

    public static void Main()
    {
        // Path where the resulting document will be saved.
        const string outputPath = "Result.docx";

        // Example token that would normally be obtained from an authentication service.
        string authToken = "valid-token";

        // URL of the external Word document. In a real case this would be a reachable endpoint.
        const string documentUrl = "https://filesamples.com/samples/document/docx/sample3.docx";

        Document doc;

        if (IsTokenValid(authToken))
        {
            // Token is valid – proceed to download the document.
            using (HttpClient httpClient = new HttpClient())
            {
                // Perform the GET request synchronously for simplicity.
                HttpResponseMessage response = httpClient.GetAsync(documentUrl).Result;
                response.EnsureSuccessStatusCode();

                // Read the content into a byte array.
                byte[] dataBytes = response.Content.ReadAsByteArrayAsync().Result;

                // Load the document from the downloaded bytes using a MemoryStream.
                using (MemoryStream stream = new MemoryStream(dataBytes))
                {
                    // Reset the stream position before loading (good practice).
                    stream.Position = 0;
                    doc = new Document(stream);
                }
            }
        }
        else
        {
            // Token is invalid – avoid the network call and create a placeholder document locally.
            doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Token validation failed. Placeholder document created.");
        }

        // Save the resulting document.
        doc.Save(outputPath);

        // Verify that the file was created successfully.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"The output file '{outputPath}' was not created.");
        }

        // Optional: indicate success (no interactive prompts required).
        Console.WriteLine("Document processing completed successfully.");
    }
}
