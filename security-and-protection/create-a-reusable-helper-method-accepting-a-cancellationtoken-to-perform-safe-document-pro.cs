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
        // Prepare a folder for temporary files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        string sourcePath = Path.Combine(artifactsDir, "Source.docx");
        string resultPath = Path.Combine(artifactsDir, "Result.docx");

        // Create a simple source document.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Original content.");
        sourceDoc.Save(sourcePath);

        // Process the document with a cancellation token (no cancellation in this demo).
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            ProcessDocument(sourcePath, resultPath, "SecretPwd", cts.Token);
        }

        // Validate that the output file exists.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Result document was not created.");

        // Verify that the document is protected as expected.
        Document resultDoc = new Document(resultPath);
        if (resultDoc.ProtectionType != ProtectionType.ReadOnly)
            throw new InvalidOperationException("Result document is not protected as expected.");
    }

    /// <summary>
    /// Loads a document, adds a paragraph, applies read‑only protection with a password,
    /// and saves the result. The operation can be cancelled via the supplied token.
    /// </summary>
    static void ProcessDocument(string inputFile, string outputFile, string password, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();

        // Load the document (no password needed for this sample).
        Document doc = new Document(inputFile);
        token.ThrowIfCancellationRequested();

        // Modify the document safely.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Additional safe text.");
        token.ThrowIfCancellationRequested();

        // Apply read‑only protection with a password.
        doc.Protect(ProtectionType.ReadOnly, password);
        token.ThrowIfCancellationRequested();

        // Save the protected document.
        doc.Save(outputFile);
        token.ThrowIfCancellationRequested();

        // Ensure the file was written.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException($"Failed to save the document to '{outputFile}'.");
    }
}
