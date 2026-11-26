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
        // Create a simple source document.
        const string sourcePath = "sample.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for background processing.");
        doc.Save(sourcePath);

        // Prepare cancellation token source.
        using var cts = new CancellationTokenSource();

        // Start background processing.
        var processingTask = ProcessDocumentAsync(sourcePath, "protected.docx", "SecretPassword", cts.Token);

        // Cancel after a short delay to demonstrate cancellation.
        _ = Task.Run(async () =>
        {
            await Task.Delay(1000);
            cts.Cancel();
        });

        try
        {
            await processingTask;
            Console.WriteLine("Document processed successfully.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Processing was cancelled.");
        }

        // Verify output file existence if processing completed.
        if (File.Exists("protected.docx"))
        {
            // Load the protected document to ensure it is indeed protected.
            var loadOptions = new LoadOptions { Password = "SecretPassword" };
            var protectedDoc = new Document("protected.docx", loadOptions);
            if (protectedDoc.ProtectionType != ProtectionType.NoProtection)
                Console.WriteLine("Protected document verified.");
            else
                throw new Exception("Protected document verification failed.");
        }
    }

    private static async Task ProcessDocumentAsync(string inputPath, string outputPath, string password, CancellationToken token)
    {
        // Simulate a long-running operation.
        await Task.Delay(3000, token);

        // Load the source document.
        var document = new Document(inputPath);

        // Apply protection.
        document.Protect(ProtectionType.ReadOnly, password);

        // Save with the same password to allow opening later.
        var saveOptions = new OoxmlSaveOptions { Password = password };
        document.Save(outputPath, saveOptions);
    }
}
