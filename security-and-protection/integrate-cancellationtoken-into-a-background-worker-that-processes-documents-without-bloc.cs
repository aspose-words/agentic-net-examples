using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static async Task Main()
    {
        // Prepare directories
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // File paths
        string protectedPath = Path.Combine(artifactsDir, "protected.docx");
        string processedPath = Path.Combine(artifactsDir, "processed.docx");

        // 1. Create a sample document and protect it with a password
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original content.");
        doc.Protect(ProtectionType.ReadOnly, "SecretPwd");
        doc.Save(protectedPath);

        // Verify the protected file exists
        if (!File.Exists(protectedPath))
            throw new InvalidOperationException("Failed to create the protected document.");

        // 2. Set up a cancellation token that cancels after 5 seconds
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        CancellationToken token = cts.Token;

        // 3. Process the document in a background task
        Task processingTask = Task.Run(() =>
        {
            // Periodically check for cancellation
            token.ThrowIfCancellationRequested();

            // Load the protected document with the correct password
            LoadOptions loadOptions = new LoadOptions("SecretPwd");
            Document loadedDoc = new Document(protectedPath, loadOptions);

            token.ThrowIfCancellationRequested();

            // Modify the document programmatically
            DocumentBuilder bg = new DocumentBuilder(loadedDoc);
            bg.Writeln("Appended text during background processing.");

            token.ThrowIfCancellationRequested();

            // Save the modified document
            loadedDoc.Save(processedPath);
        }, token);

        try
        {
            await processingTask;
            Console.WriteLine("Document processed successfully.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Document processing was canceled.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }

        // 4. Validate that the processed file exists if the task completed
        if (File.Exists(processedPath))
        {
            // Load the result to ensure it is readable
            Document resultDoc = new Document(processedPath);
            Console.WriteLine("Processed document text:");
            Console.WriteLine(resultDoc.GetText().Trim());
        }
        else
        {
            Console.WriteLine("Processed document was not created.");
        }
    }
}
