using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;

namespace AsposeWordsLoadingCancellation
{
    // Custom callback that will be invoked during document loading.
    // It throws an OperationCanceledException when the elapsed time exceeds MaxDuration.
    public class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly DateTime _loadingStartedAt;
        private const double MaxDurationSeconds = 0.0; // Immediate cancellation for demonstration.

        public LoadingProgressCallback()
        {
            _loadingStartedAt = DateTime.Now;
        }

        public void Notify(DocumentLoadingArgs args)
        {
            // Calculate elapsed time since loading began.
            DateTime now = DateTime.Now;
            double elapsedSeconds = (now - _loadingStartedAt).TotalSeconds;

            // If the elapsed time exceeds the allowed maximum, abort loading.
            if (elapsedSeconds > MaxDurationSeconds)
                throw new OperationCanceledException(
                    $"Loading canceled. EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {now}");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare a simple document to be loaded later.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);
            string sourcePath = Path.Combine(outputDir, "Sample.docx");

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document used to demonstrate loading cancellation.");
            doc.Save(sourcePath);

            // Configure LoadOptions with the custom progress callback.
            LoadOptions loadOptions = new LoadOptions
            {
                ProgressCallback = new LoadingProgressCallback()
            };

            try
            {
                // Attempt to load the document; the callback will cancel the operation.
                Document loadedDoc = new Document(sourcePath, loadOptions);
                // If loading succeeds (unlikely in this demo), indicate success.
                Console.WriteLine("Document loaded successfully (no cancellation).");
            }
            catch (OperationCanceledException ex)
            {
                // Expected path: loading was canceled by our callback.
                Console.WriteLine($"Loading was canceled: {ex.Message}");
            }
        }
    }
}
