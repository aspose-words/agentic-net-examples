using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a cancellation token that will be triggered after a short delay.
        using var cts = new CancellationTokenSource();
        cts.CancelAfter(100); // Cancel after 100 milliseconds.

        // Create a sample document with enough content to make layout processing noticeable.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Start of the document.");

        // Add many paragraphs to increase layout time.
        for (int i = 0; i < 2000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}: Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
        }

        // Attach a layout callback that monitors the cancellation token.
        doc.LayoutOptions.Callback = new LayoutCancellationCallback(cts.Token);

        // Attempt to build the layout; abort if cancellation is requested.
        try
        {
            doc.UpdatePageLayout();
        }
        catch (OperationCanceledException ex)
        {
            Console.WriteLine($"Layout building was cancelled: {ex.Message}");
        }

        // Save the document (partial layout may be present).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "LayoutCancelled.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved.");

        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Callback that aborts layout building when the cancellation token is set.
    private class LayoutCancellationCallback : IPageLayoutCallback
    {
        private readonly CancellationToken _token;

        public LayoutCancellationCallback(CancellationToken token)
        {
            _token = token;
        }

        public void Notify(PageLayoutCallbackArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException("Cancellation requested during layout building.");
        }
    }
}
