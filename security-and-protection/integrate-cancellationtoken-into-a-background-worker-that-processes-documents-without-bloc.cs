using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;   // Needed for LoadOptions
using Aspose.Words.Saving;

public class Program
{
    public static async Task Main(string[] args)
    {
        // Paths for the source and processed documents.
        string sourcePath = "sample.docx";
        string outputPath = "processed.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample document locally.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Original content.");
        // Apply read‑only protection with a password.
        sourceDoc.Protect(ProtectionType.ReadOnly, "pwd");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Set up a CancellationTokenSource (no automatic cancellation here).
        // -----------------------------------------------------------------
        using var cts = new CancellationTokenSource();
        CancellationToken token = cts.Token;

        // -----------------------------------------------------------------
        // 3. Start background processing that loads, modifies, and saves the document.
        // -----------------------------------------------------------------
        Task processingTask = Task.Run(() =>
        {
            // Throw if cancellation was requested before work starts.
            token.ThrowIfCancellationRequested();

            // Load the protected document using the correct password.
            LoadOptions loadOptions = new LoadOptions("pwd");
            Document doc = new Document(sourcePath, loadOptions);

            // Remove protection (requires the password).
            doc.Unprotect("pwd");

            // Append additional text.
            DocumentBuilder bg = new DocumentBuilder(doc);
            bg.Writeln("Processed in background.");

            // Save the modified document.
            doc.Save(outputPath);
        }, token);

        try
        {
            // Await the background task.
            await processingTask;
        }
        catch (OperationCanceledException)
        {
            // In case the operation is cancelled, report and exit.
            Console.WriteLine("Document processing was cancelled.");
            return;
        }

        // -----------------------------------------------------------------
        // 4. Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new Exception("The processed document was not saved.");

        // -----------------------------------------------------------------
        // 5. Load the result and output its text (demonstrates success).
        // -----------------------------------------------------------------
        Document resultDoc = new Document(outputPath);
        Console.WriteLine("Result document text:");
        Console.WriteLine(resultDoc.GetText());

        // Optional cleanup (commented out to keep files for inspection).
        // File.Delete(sourcePath);
        // File.Delete(outputPath);
    }
}
