using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for the demo files.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeWordsDemo");
        Directory.CreateDirectory(tempFolder);

        // Path of the sample document.
        string docPath = Path.Combine(tempFolder, "sample.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document and save it to the file system.
        // -----------------------------------------------------------------
        Document createDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(createDoc);
        builder.Writeln("This is a sample document used to demonstrate cancellation of loading.");

        // Add several pages to make the loading operation take a little longer.
        for (int i = 0; i < 50; i++)
        {
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Page {i + 2}");
        }

        createDoc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Attempt to load the document with a progress callback that
        //    aborts the operation by throwing OperationCanceledException.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions
        {
            ProgressCallback = new LoadingProgressCallback()
        };

        try
        {
            // This constructor will start loading the document and invoke the callback.
            Document loadDoc = new Document(docPath, loadOptions);
            // If loading succeeds (unlikely in this demo), we would use the document here.
            Console.WriteLine("Document loaded successfully (unexpected).");
        }
        catch (OperationCanceledException ex)
        {
            // -----------------------------------------------------------------
            // 3. Handle the cancellation and perform necessary cleanup.
            // -----------------------------------------------------------------
            Console.WriteLine($"Loading was cancelled: {ex.Message}");

            // Delete the temporary document file if it exists.
            if (File.Exists(docPath))
            {
                try
                {
                    File.Delete(docPath);
                }
                catch (Exception deleteEx)
                {
                    Console.WriteLine($"Failed to delete temporary file: {deleteEx.Message}");
                }
            }

            // Optionally delete the temporary folder if it is empty.
            try
            {
                if (Directory.Exists(tempFolder) && Directory.GetFileSystemEntries(tempFolder).Length == 0)
                {
                    Directory.Delete(tempFolder);
                }
            }
            catch (Exception dirEx)
            {
                Console.WriteLine($"Failed to delete temporary folder: {dirEx.Message}");
            }
        }
    }

    // -----------------------------------------------------------------
    // Custom loading progress callback that aborts after a short time.
    // -----------------------------------------------------------------
    private class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly DateTime _startTime = DateTime.Now;
        private const double MaxDurationSeconds = 0.1; // Cancel after 100 ms.

        public void Notify(DocumentLoadingArgs args)
        {
            double elapsed = (DateTime.Now - _startTime).TotalSeconds;
            if (elapsed > MaxDurationSeconds)
            {
                throw new OperationCanceledException(
                    $"Loading cancelled after {elapsed:F2}s (EstimatedProgress = {args.EstimatedProgress}).");
            }
        }
    }
}
