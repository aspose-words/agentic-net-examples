using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        for (int i = 0; i < 1000; i++)
        {
            builder.Writeln($"Paragraph {i}");
        }

        // Set up a cancellation token that will be triggered shortly.
        using var cts = new CancellationTokenSource();
        cts.CancelAfter(10); // Cancel after 10 milliseconds.

        // Attach a layout callback that monitors the token.
        doc.LayoutOptions.Callback = new CancelLayoutCallback(cts.Token);

        try
        {
            // Build the page layout. The callback will abort if cancellation is requested.
            doc.UpdatePageLayout();
            Console.WriteLine("Layout completed successfully.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Layout was cancelled via token.");
        }

        // Save the document to verify that saving still works.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved as expected.");
    }

    // Callback implementation that checks the cancellation token during layout.
    private class CancelLayoutCallback : IPageLayoutCallback
    {
        private readonly CancellationToken _token;

        public CancelLayoutCallback(CancellationToken token) => _token = token;

        public void Notify(PageLayoutCallbackArgs args)
        {
            // The WatchDog event is a frequent checkpoint suitable for abort checks.
            if (args.Event == PageLayoutEvent.WatchDog && _token.IsCancellationRequested)
                throw new OperationCanceledException("Cancellation requested during layout processing.");
        }
    }
}
