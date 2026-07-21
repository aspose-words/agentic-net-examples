using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "CancelledLayout.docx");

        // Create a sample document with enough content to make layout processing noticeable.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 200; i++)
        {
            builder.Writeln($"Paragraph {i + 1}: The quick brown fox jumps over the lazy dog.");
        }

        // Set up a cancellation token that will be triggered after a short delay.
        using var cts = new CancellationTokenSource();
        Task.Run(() =>
        {
            Thread.Sleep(100); // Simulate external request to cancel.
            cts.Cancel();
        });

        // Attach a layout callback that monitors the token.
        doc.LayoutOptions.Callback = new CancelableLayoutCallback(cts.Token);

        try
        {
            // Start layout building. The callback will abort if cancellation is requested.
            doc.UpdatePageLayout();
        }
        catch (OperationCanceledException)
        {
            // Layout was cancelled; we can continue with whatever partial layout exists.
            Console.WriteLine("Layout operation was cancelled.");
        }

        // Save the document (partial layout is still a valid document).
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to save the document.");

        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Callback that aborts layout when the cancellation token is set.
    private class CancelableLayoutCallback : IPageLayoutCallback
    {
        private readonly CancellationToken _token;

        public CancelableLayoutCallback(CancellationToken token)
        {
            _token = token;
        }

        public void Notify(PageLayoutCallbackArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException("Layout cancelled via token.");
        }
    }
}
