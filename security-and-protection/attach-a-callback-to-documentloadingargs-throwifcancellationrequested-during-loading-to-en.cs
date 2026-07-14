using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using System.Threading;

public class Program
{
    // Callback that aborts loading by throwing an OperationCanceledException.
    private class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly DateTime _startTime = DateTime.Now;
        private const double MaxDurationSeconds = 0.1; // short duration to trigger cancellation

        public void Notify(DocumentLoadingArgs args)
        {
            // If the loading takes longer than the allowed time, request cancellation.
            double elapsed = (DateTime.Now - _startTime).TotalSeconds;
            if (elapsed > MaxDurationSeconds)
            {
                // Some versions expose ThrowIfCancellationRequested; call it if available.
                // args.ThrowIfCancellationRequested(); // Uncomment if the method exists.
                throw new OperationCanceledException(
                    $"Loading cancelled after {elapsed:F2}s (EstimatedProgress = {args.EstimatedProgress}).");
            }
        }
    }

    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a simple document and save it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");
        string filePath = Path.Combine(outputDir, "Sample.docx");
        doc.Save(filePath);

        // 2. Load the document with a progress callback that may cancel the operation.
        LoadOptions loadOptions = new LoadOptions
        {
            ProgressCallback = new LoadingProgressCallback()
        };

        try
        {
            Document loadedDoc = new Document(filePath, loadOptions);
            Console.WriteLine("Document loaded successfully.");
        }
        catch (OperationCanceledException ex)
        {
            Console.WriteLine($"Loading was cancelled: {ex.Message}");
        }

        // 3. Load the document without cancellation to demonstrate normal loading.
        Document normalLoad = new Document(filePath);
        Console.WriteLine($"Normal load text: {normalLoad.GetText().Trim()}");
    }
}
