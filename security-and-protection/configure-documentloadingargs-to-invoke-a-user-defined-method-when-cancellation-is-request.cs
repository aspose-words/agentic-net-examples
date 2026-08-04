using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Create a simple document and save it locally.
        const string fileName = "Sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");
        doc.Save(fileName);

        // Configure load options with a progress callback that cancels loading.
        LoadOptions loadOptions = new LoadOptions
        {
            ProgressCallback = new LoadingProgressCallback()
        };

        try
        {
            // Attempt to load the document; the callback will abort the operation.
            Document loadedDoc = new Document(fileName, loadOptions);
            Console.WriteLine("Document loaded successfully.");
        }
        catch (OperationCanceledException ex)
        {
            // Expected path when the callback triggers cancellation.
            Console.WriteLine($"Loading canceled: {ex.Message}");
        }
        finally
        {
            // Clean up the temporary file.
            if (File.Exists(fileName))
                File.Delete(fileName);
        }
    }

    // User‑defined callback that aborts loading after a minimal elapsed time.
    private class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly DateTime _loadingStartedAt;
        private const double MaxDuration = 0.0; // Cancel immediately.

        public LoadingProgressCallback()
        {
            _loadingStartedAt = DateTime.Now;
        }

        public void Notify(DocumentLoadingArgs args)
        {
            double elapsedSeconds = (DateTime.Now - _loadingStartedAt).TotalSeconds;
            if (elapsedSeconds > MaxDuration)
                throw new OperationCanceledException(
                    $"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
        }
    }
}
