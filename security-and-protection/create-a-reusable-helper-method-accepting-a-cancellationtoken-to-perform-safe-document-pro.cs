using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    // Reusable helper that processes a document safely with cancellation support.
    public static void ProcessDocument(string inputPath, string outputPath, string password, CancellationToken token)
    {
        // Check for cancellation before starting.
        token.ThrowIfCancellationRequested();

        // Load the document with the provided password (if any).
        var loadOptions = new LoadOptions
        {
            Password = password
        };
        var doc = new Document(inputPath, loadOptions);

        token.ThrowIfCancellationRequested();

        // Apply read‑only protection using the same password.
        doc.Protect(ProtectionType.ReadOnly, password);

        token.ThrowIfCancellationRequested();

        // Save the protected document with encryption.
        var saveOptions = new OoxmlSaveOptions
        {
            Password = password
        };
        doc.Save(outputPath, saveOptions);

        token.ThrowIfCancellationRequested();

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the protected document at '{outputPath}'.");
    }

    public static void Main()
    {
        // Prepare a temporary folder for the demo files.
        string demoFolder = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(demoFolder);

        // Create a simple source document.
        string sourcePath = Path.Combine(demoFolder, "source.docx");
        var sourceDoc = new Document();
        var paragraph = new Paragraph(sourceDoc);
        paragraph.AppendChild(new Run(sourceDoc, "Hello, Aspose.Words!"));
        sourceDoc.FirstSection.Body.AppendChild(paragraph);
        sourceDoc.Save(sourcePath);

        // Define the output path and password.
        string protectedPath = Path.Combine(demoFolder, "protected.docx");
        string password = "Secret123";

        // Process the document with a non‑canceled token.
        var cancellationToken = CancellationToken.None;
        ProcessDocument(sourcePath, protectedPath, password, cancellationToken);

        // Indicate success.
        Console.WriteLine("Document processed and protected successfully.");
    }
}
