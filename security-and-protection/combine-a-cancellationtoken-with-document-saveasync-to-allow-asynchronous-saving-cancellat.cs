using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static async Task Main()
    {
        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document will be saved asynchronously with cancellation support.");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AsyncCancelled.docx");

        // Prepare a cancellation token that will be cancelled after a short delay.
        using var cts = new CancellationTokenSource();
        cts.CancelAfter(100); // Cancel after 100 ms to simulate a quick abort.
        CancellationToken token = cts.Token;

        try
        {
            // Perform the save operation on a background thread and observe the token.
            await Task.Run(() =>
            {
                // Throw if cancellation was already requested.
                token.ThrowIfCancellationRequested();

                // Save the document using normal synchronous API.
                // The token is not directly passed because Aspose.Words does not expose SaveAsync.
                doc.Save(outputPath, new OoxmlSaveOptions(SaveFormat.Docx));
            }, token);

            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Document saving was cancelled.");
        }

        // Verify whether the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Output file exists.");
        }
        else
        {
            Console.WriteLine("Output file does not exist.");
        }
    }
}
