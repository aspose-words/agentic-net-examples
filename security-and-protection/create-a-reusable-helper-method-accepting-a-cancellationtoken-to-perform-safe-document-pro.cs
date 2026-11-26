using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a cancellation token source (no cancellation in this example)
        using var cts = new CancellationTokenSource();

        // Process the document safely
        ProcessDocumentSafely("Sample.docx", "Protected.docx", "myPassword", cts.Token);

        // Verify that the protected file was created
        if (!File.Exists("Protected.docx"))
            throw new Exception("Protected document was not created.");

        Console.WriteLine("Document processed and protected successfully.");
    }

    /// <summary>
    /// Creates a simple document, applies read‑only protection with a password,
    /// saves it, and validates the protection while respecting the provided CancellationToken.
    /// </summary>
    /// <param name="sourcePath">Path to temporarily save the unprotected document.</param>
    /// <param name="outputPath">Path to save the protected document.</param>
    /// <param name="password">Password used for protection and encryption.</param>
    /// <param name="token">Cancellation token to allow safe abort.</param>
    private static void ProcessDocumentSafely(string sourcePath, string outputPath, string password, CancellationToken token)
    {
        // Check for cancellation before starting
        token.ThrowIfCancellationRequested();

        // Create a new document with a single line of text
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");

        // Optional: save the unprotected version (demonstrates load/save workflow)
        token.ThrowIfCancellationRequested();
        doc.Save(sourcePath);

        // Apply read‑only protection with the specified password
        token.ThrowIfCancellationRequested();
        doc.Protect(ProtectionType.ReadOnly, password);

        // Save the protected document using OoxmlSaveOptions with password encryption
        token.ThrowIfCancellationRequested();
        var saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Password = password
        };
        doc.Save(outputPath, saveOptions);

        // Reopen the protected document to verify that protection was applied
        token.ThrowIfCancellationRequested();
        var loadOptions = new LoadOptions { Password = password };
        var protectedDoc = new Document(outputPath, loadOptions);

        // Validate protection type
        token.ThrowIfCancellationRequested();
        if (protectedDoc.ProtectionType != ProtectionType.ReadOnly)
            throw new Exception("Document protection was not applied correctly.");
    }
}
