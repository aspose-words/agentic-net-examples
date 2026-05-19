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
        string originalPath = Path.Combine(artifactsDir, "Original.docx");
        string protectedPath = Path.Combine(artifactsDir, "Protected.docx");
        string processedPath = Path.Combine(artifactsDir, "Processed.docx");

        // 1. Create a simple document
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("Hello world! This is the original document.");
        originalDoc.Save(originalPath);

        // 2. Protect the document with a password
        originalDoc.Protect(ProtectionType.ReadOnly, "SecretPwd");
        originalDoc.Save(protectedPath);

        // 3. Set up a cancellation token that cancels after 5 seconds
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        CancellationToken token = cts.Token;

        try
        {
            // 4. Process the protected document in a background task
            await Task.Run(() =>
            {
                // Periodically check for cancellation
                token.ThrowIfCancellationRequested();

                // Load the protected document with the correct password
                LoadOptions loadOptions = new LoadOptions("SecretPwd");
                Document protectedDoc = new Document(protectedPath, loadOptions);

                token.ThrowIfCancellationRequested();

                // Modify the document (add a line)
                DocumentBuilder procBuilder = new DocumentBuilder(protectedDoc);
                procBuilder.Writeln("This line was added in the background worker.");

                token.ThrowIfCancellationRequested();

                // Save the processed document
                protectedDoc.Save(processedPath);
            }, token);
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Document processing was canceled.");
        }

        // 5. Validate that the processed file exists
        if (File.Exists(processedPath))
        {
            Console.WriteLine($"Processed document saved to: {processedPath}");
        }
        else
        {
            Console.WriteLine("Processed document was not created.");
        }
    }
}
