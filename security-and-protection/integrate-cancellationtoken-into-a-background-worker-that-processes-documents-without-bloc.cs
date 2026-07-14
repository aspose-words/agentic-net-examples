using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static async Task Main(string[] args)
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");
        string samplePath = Path.Combine(outputDir, "sample.docx");
        doc.Save(samplePath);

        // 2. Protect the document with a password.
        doc.Protect(ProtectionType.ReadOnly, "SecretPwd");
        string protectedPath = Path.Combine(outputDir, "protected.docx");
        doc.Save(protectedPath);

        // 3. Set up a cancellation token that will cancel after 3 seconds.
        using var cts = new CancellationTokenSource();
        cts.CancelAfter(TimeSpan.FromSeconds(3));
        CancellationToken token = cts.Token;

        // 4. Run document processing in the background.
        string processedPath = Path.Combine(outputDir, "processed.docx");
        Task processingTask = Task.Run(() => ProcessDocumentAsync(protectedPath, processedPath, token), token);

        try
        {
            await processingTask;
            Console.WriteLine("Document processing completed.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Document processing was canceled.");
        }

        // 5. Validate that the processed file exists (if not canceled).
        if (File.Exists(processedPath))
        {
            Console.WriteLine($"Processed file saved at: {processedPath}");
        }
        else
        {
            Console.WriteLine("Processed file was not created.");
        }
    }

    private static async Task ProcessDocumentAsync(string inputPath, string outputPath, CancellationToken token)
    {
        // Simulate some initial delay.
        await Task.Delay(500, token);

        // Load the protected document with the correct password.
        LoadOptions loadOptions = new LoadOptions("SecretPwd");
        Document protectedDoc = new Document(inputPath, loadOptions);

        // Periodically check for cancellation.
        token.ThrowIfCancellationRequested();

        // Unprotect the document (password not required for Unprotect()).
        protectedDoc.Unprotect();

        // Add additional content.
        DocumentBuilder builder = new DocumentBuilder(protectedDoc);
        builder.Writeln("Additional text added in background processing.");

        // Simulate longer work.
        for (int i = 0; i < 5; i++)
        {
            token.ThrowIfCancellationRequested();
            await Task.Delay(300, token); // pretend work
        }

        // Save the modified document.
        protectedDoc.Save(outputPath);
    }
}
