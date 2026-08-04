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
        // Create a sample document with enough content to make layout processing noticeable.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");
        for (int i = 0; i < 1000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        // Set up a cancellation token that will be triggered shortly after layout starts.
        CancellationTokenSource cts = new CancellationTokenSource();
        Task.Run(async () =>
        {
            await Task.Delay(10); // Adjust delay as needed to simulate cancellation during layout.
            cts.Cancel();
        });

        // Attach a layout callback that checks the token and aborts if cancellation is requested.
        doc.LayoutOptions.Callback = new CancelableLayoutCallback(cts.Token);

        try
        {
            // Begin layout building. The callback will throw if cancellation occurs.
            doc.UpdatePageLayout();

            // If layout completes, save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to: {outputPath}");
        }
        catch (OperationCanceledException)
        {
            // Layout was aborted due to cancellation.
            Console.WriteLine("Layout building was canceled.");
        }
    }

    // Callback implementation that monitors the cancellation token.
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
            {
                // Throwing an OperationCanceledException aborts the layout process.
                throw new OperationCanceledException("Cancellation requested during layout.");
            }
        }
    }
}
