using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;

public class LoadingProgressCallback : IDocumentLoadingCallback
{
    private readonly DateTime _startTime;
    private const double MaxDurationSeconds = 0.1; // Cancel quickly for demo

    public LoadingProgressCallback()
    {
        _startTime = DateTime.Now;
    }

    public void Notify(DocumentLoadingArgs args)
    {
        // Cancel loading if it exceeds the allowed duration.
        if ((DateTime.Now - _startTime).TotalSeconds > MaxDurationSeconds)
            throw new OperationCanceledException(
                $"Loading canceled after {args.EstimatedProgress}% progress.");
    }
}

public class Program
{
    public static void Main()
    {
        // Path for the temporary document.
        string filePath = Path.Combine(Path.GetTempPath(), "Sample.docx");

        // Create a simple document and save it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        doc.Save(filePath);

        // Prepare load options with a progress callback that will cancel the load.
        LoadOptions loadOptions = new LoadOptions
        {
            ProgressCallback = new LoadingProgressCallback()
        };

        try
        {
            // Attempt to load the document; this will be canceled.
            Document loadedDoc = new Document(filePath, loadOptions);
            // If loading succeeds (unlikely), write a message.
            Console.WriteLine("Document loaded successfully (unexpected).");
        }
        catch (OperationCanceledException ex)
        {
            // Handle the cancellation and perform cleanup.
            Console.WriteLine($"OperationCanceledException caught: {ex.Message}");
            // Delete the temporary file as part of cleanup.
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    Console.WriteLine("Temporary file deleted.");
                }
            }
            catch (Exception cleanupEx)
            {
                Console.WriteLine($"Cleanup error: {cleanupEx.Message}");
            }
        }
        finally
        {
            // Ensure any remaining resources are released.
            // In this simple example there are no additional resources to dispose.
        }
    }
}
