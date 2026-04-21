using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Path for the output document.
        const string outputPath = "LayoutCancelled.docx";

        // Create a sample document with a lot of content to ensure layout takes some time.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Start of the document.");
        for (int i = 0; i < 2000; i++)
        {
            builder.Writeln($"Line {i + 1}");
        }
        builder.Writeln("End of the document.");

        // Set up a cancellation token that will be triggered shortly.
        using var cts = new CancellationTokenSource();
        cts.CancelAfter(10); // Cancel after 10 milliseconds.

        // Assign a layout callback that checks the token and aborts if cancellation is requested.
        doc.LayoutOptions.Callback = new LayoutCancellationCallback(cts.Token);

        try
        {
            // Build the page layout. The callback may throw OperationCanceledException.
            doc.UpdatePageLayout();

            // If layout finishes, save the document.
            doc.Save(outputPath);
            Console.WriteLine("Layout completed successfully and document saved.");
        }
        catch (OperationCanceledException ex)
        {
            // Layout was aborted. Save whatever state the document is in.
            Console.WriteLine($"Layout aborted: {ex.Message}");
            doc.Save(outputPath);
            Console.WriteLine("Partial document saved after cancellation.");
        }

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output document was not created.");
        }
    }

    // Callback implementation that monitors the cancellation token.
    private class LayoutCancellationCallback : IPageLayoutCallback
    {
        private readonly CancellationToken _token;

        public LayoutCancellationCallback(CancellationToken token)
        {
            _token = token;
        }

        public void Notify(PageLayoutCallbackArgs args)
        {
            // Abort the layout process as soon as cancellation is requested.
            if (_token.IsCancellationRequested)
            {
                throw new OperationCanceledException("Layout building was cancelled via token.");
            }
        }
    }
}
