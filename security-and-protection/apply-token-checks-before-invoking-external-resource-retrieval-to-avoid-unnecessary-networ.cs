using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Expected token for accessing the external resource.
        const string expectedToken = "secret123";

        // Simulated token received from a caller or configuration.
        string providedToken = "secret123";

        // Verify token before making any network request.
        if (providedToken != expectedToken)
        {
            Console.WriteLine("Invalid token. Skipping external document retrieval.");
            return;
        }

        // Token is valid – proceed to download the document.
        const string documentUrl = "https://github.com/aspose-words/Aspose.Words-for-.NET/raw/master/Examples/Data/Document.docx";

        using var httpClient = new HttpClient();
        byte[] documentBytes = httpClient.GetByteArrayAsync(documentUrl).Result;

        // Load the downloaded bytes into a MemoryStream for Aspose.Words.
        using var memoryStream = new MemoryStream(documentBytes);
        memoryStream.Position = 0;

        // Create a Document object from the stream.
        var document = new Document(memoryStream);

        // Apply read‑only protection with a password.
        document.Protect(ProtectionType.ReadOnly, "password");

        // Save the protected document locally.
        const string outputPath = "ProtectedOutput.docx";
        document.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The protected document was not saved as expected.");
        }

        Console.WriteLine($"Document downloaded, protected, and saved to '{outputPath}'.");
    }
}
