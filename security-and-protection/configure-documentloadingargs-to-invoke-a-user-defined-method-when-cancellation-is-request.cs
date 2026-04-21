using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class LoadingProgressCallback : IDocumentLoadingCallback
{
    private readonly DateTime _loadingStartedAt;
    private const double MaxDuration = 0.0; // seconds, cancel immediately

    public LoadingProgressCallback()
    {
        _loadingStartedAt = DateTime.Now;
    }

    public void Notify(DocumentLoadingArgs args)
    {
        double elapsedSeconds = (DateTime.Now - _loadingStartedAt).TotalSeconds;
        if (elapsedSeconds > MaxDuration)
            throw new OperationCanceledException($"Loading canceled. EstimatedProgress = {args.EstimatedProgress}");
    }
}

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for the sample files.
        string tempDir = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(tempDir);

        // Create a simple document and save it.
        string sourcePath = Path.Combine(tempDir, "Sample.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");
        doc.Save(sourcePath);

        // Verify that the document was saved.
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException("Failed to create the source document.", sourcePath);

        // Set up load options with a progress callback that cancels loading.
        LoadOptions loadOptions = new LoadOptions
        {
            ProgressCallback = new LoadingProgressCallback()
        };

        try
        {
            // Attempt to load the document; cancellation is expected.
            Document loadedDoc = new Document(sourcePath, loadOptions);
            // If loading succeeds (unlikely), indicate that cancellation did not occur.
            Console.WriteLine("Document loaded without cancellation.");
        }
        catch (OperationCanceledException ex)
        {
            // Expected path: loading was canceled by our callback.
            Console.WriteLine($"OperationCanceledException caught: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Any other exception is unexpected.
            Console.WriteLine($"Unexpected exception: {ex.GetType().Name} - {ex.Message}");
        }
    }
}
