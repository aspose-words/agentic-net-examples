using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Entry point
    public static async Task Main()
    {
        // Prepare sample source document
        const string sourcePath = "source.docx";
        const string protectedPath = "protected.docx";
        const string protectionPassword = "Secret123";

        CreateSampleDocument(sourcePath);

        // Set up cancellation token (not cancelled in this example)
        using var cts = new CancellationTokenSource();
        CancellationToken token = cts.Token;

        // Run document processing in background
        await Task.Run(() => ProcessDocument(sourcePath, protectedPath, protectionPassword, token), token);

        // Verify output exists
        if (!File.Exists(protectedPath))
            throw new InvalidOperationException($"Failed to create protected document at '{protectedPath}'.");

        // Clean up sample files (optional)
        File.Delete(sourcePath);
        File.Delete(protectedPath);
    }

    // Creates a simple DOCX file with some text
    private static void CreateSampleDocument(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for protection demonstration.");
        doc.Save(path);
    }

    // Loads a document, applies read‑only protection with a password, and saves it
    private static void ProcessDocument(string inputPath, string outputPath, string password, CancellationToken token)
    {
        // Throw if cancellation requested before starting
        token.ThrowIfCancellationRequested();

        // Load the source document (no password needed for the source)
        var doc = new Document(inputPath);

        // Apply read‑only protection with the provided password
        doc.Protect(ProtectionType.ReadOnly, password);

        // Save the protected document
        var saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        doc.Save(outputPath, saveOptions);

        // Throw if cancellation requested after processing
        token.ThrowIfCancellationRequested();
    }
}
