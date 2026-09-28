using System;
using System.IO;
using System.Threading;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");

        // Use CancellationTokenSource within a using block to guarantee disposal.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Perform processing that could be cancelled.
            ProcessDocument(doc, cts.Token);
        } // cts is disposed here, freeing system resources.

        // Save the processed document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved as expected.");
    }

    private static void ProcessDocument(Document doc, CancellationToken token)
    {
        // Apply read‑only protection with a password.
        doc.Protect(ProtectionType.ReadOnly, "password");

        // Check for cancellation request.
        if (token.IsCancellationRequested)
            token.ThrowIfCancellationRequested();

        // Placeholder for additional processing logic.
    }
}
