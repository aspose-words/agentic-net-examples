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
        // Define output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ProtectedDocument.docx");

        // Ensure any previous file is removed.
        if (File.Exists(outputPath))
            File.Delete(outputPath);

        // Create a CancellationTokenSource that will be disposed after use.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Simulate a cancellation token being passed to a long‑running operation.
            // In this example the token is not actually used to cancel, but the pattern shows proper disposal.
            ProcessDocument(cts.Token, outputPath);
        }

        // Verify that the document was saved.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"The expected output file was not created: {outputPath}");

        // Load the saved document to confirm it is protected.
        LoadOptions loadOptions = new LoadOptions(); // No password needed for protection type.
        Document loadedDoc = new Document(outputPath, loadOptions);

        if (loadedDoc.ProtectionType != ProtectionType.ReadOnly)
            throw new InvalidOperationException("The document is not protected as expected.");
    }

    private static void ProcessDocument(CancellationToken token, string outputPath)
    {
        // Create a blank document.
        Document doc = new Document();

        // Add some content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document is protected with a read‑only restriction.");

        // Apply protection (no password needed for this demo).
        doc.Protect(ProtectionType.ReadOnly);

        // Save the protected document.
        // The token is not required by Aspose.Words, but we keep the method signature consistent.
        doc.Save(outputPath);
    }
}
