using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;

public class Program
{
    public static async Task Main(string[] args)
    {
        const string filePath = "sample.docx";

        // Create a simple document and save it.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for cancellation token demo.");
        doc.Save(filePath);

        if (!File.Exists(filePath))
            throw new FileNotFoundException("Failed to create the sample document.", filePath);

        // Set up a cancellation token source that cancels after a short delay.
        using var cts = new CancellationTokenSource();

        // Cancel after 100 ms.
        _ = Task.Run(async () =>
        {
            await Task.Delay(100);
            cts.Cancel();
        });

        try
        {
            // Load the document inside a Task that observes the cancellation token.
            var loadedDoc = await Task.Run(() => new Document(filePath), cts.Token);
            Console.WriteLine("Document loaded successfully.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Document loading was canceled.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
