using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for the demo files.
        string artifactsDir = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample Word document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Hello world! This is the original document.");

        // Apply read‑only protection with a password.
        sourceDoc.Protect(ProtectionType.ReadOnly, "SecretPwd");

        // Save the protected document.
        string sourcePath = Path.Combine(artifactsDir, "ProtectedSource.docx");
        sourceDoc.Save(sourcePath);

        // Verify that the source file was created.
        if (!File.Exists(sourcePath))
            throw new InvalidOperationException("Failed to create the source document.");

        // Define the output path for the processed document.
        string outputPath = Path.Combine(artifactsDir, "Processed.docx");

        // Set up a cancellation token that will cancel after 3 seconds.
        using var cts = new CancellationTokenSource();
        cts.CancelAfter(TimeSpan.FromSeconds(3));
        CancellationToken token = cts.Token;

        // Start the background processing task.
        Task processingTask = Task.Run(() => ProcessDocumentAsync(sourcePath, outputPath, token), token);

        try
        {
            // Wait for the task to complete (or be cancelled).
            processingTask.Wait(token);
        }
        catch (AggregateException ae)
        {
            // Unwrap the cancellation exception if the operation was cancelled.
            if (ae.InnerException is OperationCanceledException)
                Console.WriteLine("Document processing was cancelled.");
            else
                throw;
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Document processing was cancelled.");
        }

        // If the task completed without cancellation, verify the output file.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Processed document saved to: {outputPath}");
        }
        else
        {
            Console.WriteLine("Processed document was not created (likely cancelled).");
        }
    }

    // Background worker that loads, modifies, and saves the document.
    private static void ProcessDocumentAsync(string inputPath, string outputPath, CancellationToken token)
    {
        // Simulate some preparatory work that can be cancelled.
        for (int i = 0; i < 5; i++)
        {
            token.ThrowIfCancellationRequested();
            Thread.Sleep(500); // Simulate work.
        }

        // Load the protected document. Protection does not encrypt the file,
        // so it can be opened without a password.
        Document doc = new Document(inputPath, new LoadOptions());

        // Check for cancellation before modifying the document.
        token.ThrowIfCancellationRequested();

        // Append a new paragraph to indicate processing.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document processed by background worker.");

        // Save the modified document. The .docx extension determines the format,
        // so we can use the simple overload without specifying SaveOptions.
        doc.Save(outputPath);

        // Final cancellation check (optional).
        token.ThrowIfCancellationRequested();
    }
}
