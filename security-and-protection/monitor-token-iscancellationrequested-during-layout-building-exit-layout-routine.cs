using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Layout;

namespace LayoutCancellationDemo
{
    // Custom callback that checks a CancellationToken during layout building.
    public class CancellationPageLayoutCallback : IPageLayoutCallback
    {
        private readonly CancellationToken _cancellationToken;

        public CancellationPageLayoutCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
        }

        // This method is called repeatedly during layout building.
        // Throwing an exception aborts the layout process.
        public void Notify(PageLayoutCallbackArgs args)
        {
            // The WatchDog event is a frequent checkpoint suitable for cancellation checks.
            if (args.Event == PageLayoutEvent.WatchDog && _cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException("Layout building was cancelled.");
        }
    }

    class Program
    {
        static void Main()
        {
            // Prepare a cancellation token that will be triggered after a short delay.
            using var cts = new CancellationTokenSource();
            cts.CancelAfter(TimeSpan.FromSeconds(1)); // Adjust as needed.

            // Create a sample document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("First page content.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Second page content.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Third page content.");

            // Assign the custom layout callback.
            doc.LayoutOptions.Callback = new CancellationPageLayoutCallback(cts.Token);

            try
            {
                // Start building the page layout. The callback will abort if cancellation is requested.
                doc.UpdatePageLayout();
                Console.WriteLine("Layout completed successfully.");
            }
            catch (OperationCanceledException ex)
            {
                Console.WriteLine($"Layout aborted: {ex.Message}");
            }

            // Optionally save the partially built document (if any layout was completed).
            // The save operation will use the layout that was built up to the point of cancellation.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "CancelledLayout.docx");
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
