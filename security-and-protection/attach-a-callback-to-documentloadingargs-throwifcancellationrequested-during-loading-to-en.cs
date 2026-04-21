using System;
using System.Diagnostics;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class LoadingProgressCallback : IDocumentLoadingCallback
{
    private readonly Stopwatch _stopwatch = Stopwatch.StartNew();
    private readonly double _maxSeconds;

    public LoadingProgressCallback(double maxSeconds)
    {
        _maxSeconds = maxSeconds;
    }

    public void Notify(DocumentLoadingArgs args)
    {
        // Cancel loading if the elapsed time exceeds the allowed limit.
        if (_stopwatch.Elapsed.TotalSeconds > _maxSeconds)
        {
            // Throw an OperationCanceledException to abort the loading process.
            throw new OperationCanceledException(
                $"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
        }
    }
}

public class Program
{
    public static void Main()
    {
        // Prepare a sample document.
        string fileName = Path.Combine(Directory.GetCurrentDirectory(), "sample.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");
        doc.Save(fileName);

        // Verify that the file was created.
        if (!File.Exists(fileName))
            throw new InvalidOperationException("Failed to create the sample document.");

        // Attempt to load the document with a callback that cancels immediately.
        LoadOptions cancelOptions = new LoadOptions
        {
            ProgressCallback = new LoadingProgressCallback(0) // 0 seconds -> immediate cancellation
        };

        try
        {
            Document canceledDoc = new Document(fileName, cancelOptions);
            // If loading succeeds, this line will not be reached.
            Console.WriteLine("Unexpectedly loaded the document without cancellation.");
        }
        catch (OperationCanceledException ex)
        {
            Console.WriteLine("Loading was cancelled as expected: " + ex.Message);
        }

        // Load the document again with a non‑cancelling callback.
        LoadOptions normalOptions = new LoadOptions
        {
            ProgressCallback = new LoadingProgressCallback(5) // 5 seconds allowed
        };

        Document loadedDoc = new Document(fileName, normalOptions);
        Console.WriteLine("Document loaded successfully. Text content:");
        Console.WriteLine(loadedDoc.GetText().Trim());
    }
}
