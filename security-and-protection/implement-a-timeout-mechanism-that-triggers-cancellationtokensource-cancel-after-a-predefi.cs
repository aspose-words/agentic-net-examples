using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static async Task Main()
    {
        // Prepare a folder for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string docPath = Path.Combine(artifactsDir, "Sample.docx");

        // Create a simple Word document and save it locally.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        doc.Save(docPath);

        // Set up a cancellation token that will be cancelled after 2 seconds.
        using var cts = new CancellationTokenSource();
        var timer = new System.Timers.Timer(2000) { AutoReset = false };
        timer.Elapsed += (s, e) => cts.Cancel();
        timer.Start();

        try
        {
            // Attempt to load the document with a simulated long‑running operation.
            await LoadDocumentWithTimeoutAsync(docPath, cts.Token);
            Console.WriteLine("Document loaded successfully before timeout.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Loading was cancelled due to timeout.");
        }
        finally
        {
            timer.Dispose();
        }
    }

    private static async Task LoadDocumentWithTimeoutAsync(string path, CancellationToken token)
    {
        // Simulate a lengthy operation that periodically checks the cancellation token.
        await Task.Run(() =>
        {
            for (int i = 0; i < 10; i++)
            {
                token.ThrowIfCancellationRequested();
                // Simulate work (e.g., processing a large document).
                Thread.Sleep(500);
            }

            // Actual document loading (fast for this small sample).
            var loadOptions = new LoadOptions(); // No password required.
            Document loaded = new Document(path, loadOptions);
        }, token);
    }
}
