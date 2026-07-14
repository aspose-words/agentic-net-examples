using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

class LoadingProgressCallback : IDocumentLoadingCallback
{
    // Time when loading started.
    private readonly DateTime _loadingStartedAt;
    // Set to zero to trigger cancellation immediately.
    private const double MaxDuration = 0.0;

    public LoadingProgressCallback()
    {
        _loadingStartedAt = DateTime.Now;
    }

    // This method is called repeatedly during document loading.
    public void Notify(DocumentLoadingArgs args)
    {
        double elapsedSeconds = (DateTime.Now - _loadingStartedAt).TotalSeconds;
        if (elapsedSeconds > MaxDuration)
            throw new OperationCanceledException($"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
    }
}

public class Program
{
    public static void Main()
    {
        // Path for the temporary document.
        string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");

        // Create a simple document and save it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");
        doc.Save(filePath);

        // Configure load options with the custom progress callback.
        LoadOptions loadOptions = new LoadOptions { ProgressCallback = new LoadingProgressCallback() };

        try
        {
            // Attempt to load the document; the callback will cancel the operation.
            Document loadedDoc = new Document(filePath, loadOptions);
            Console.WriteLine("Document loaded successfully (this line should not appear).");
        }
        catch (OperationCanceledException ex)
        {
            // Expected outcome when the callback cancels loading.
            Console.WriteLine($"Loading was canceled: {ex.Message}");
        }

        // Clean up the temporary file.
        if (File.Exists(filePath))
            File.Delete(filePath);
    }
}
