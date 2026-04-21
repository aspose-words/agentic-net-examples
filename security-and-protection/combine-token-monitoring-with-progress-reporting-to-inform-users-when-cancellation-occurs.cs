using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Prepare paths for the sample documents.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string originalPath = Path.Combine(artifactsDir, "Original.docx");
        string protectedPath = Path.Combine(artifactsDir, "Protected.docx");

        // 1. Create a simple document and save it normally.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document used to demonstrate token monitoring and progress reporting.");
        doc.Save(originalPath);
        if (!File.Exists(originalPath))
            throw new InvalidOperationException("Failed to create the original document.");

        // 2. Attempt to load the document with a loading progress callback that cancels quickly.
        try
        {
            LoadOptions loadOptions = new LoadOptions
            {
                ProgressCallback = new LoadingProgressCallback()
            };
            Document loadedDoc = new Document(originalPath, loadOptions);
            // If loading succeeds (unlikely), delete the document to keep the example clean.
            loadedDoc.Save(protectedPath);
        }
        catch (OperationCanceledException ex)
        {
            Console.WriteLine($"Loading was cancelled: {ex.Message}");
        }

        // 3. Attempt to save the original document with a saving progress callback that cancels quickly.
        try
        {
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                ProgressCallback = new SavingProgressCallback()
            };
            doc.Save(protectedPath, saveOptions);
        }
        catch (OperationCanceledException ex)
        {
            Console.WriteLine($"Saving was cancelled: {ex.Message}");
        }

        // 4. Verify that the protected file was not created due to cancellation.
        if (File.Exists(protectedPath))
            Console.WriteLine("Protected document was saved (cancellation did not occur).");
        else
            Console.WriteLine("Protected document was not saved because the operation was cancelled.");
    }

    // Loading progress callback that aborts after a very short duration.
    private class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly DateTime _start = DateTime.Now;
        private const double MaxDurationSeconds = 0.001; // Very short to force cancellation.

        public void Notify(DocumentLoadingArgs args)
        {
            double elapsed = (DateTime.Now - _start).TotalSeconds;
            if (elapsed > MaxDurationSeconds)
                throw new OperationCanceledException($"Loading cancelled. EstimatedProgress = {args.EstimatedProgress}");
        }
    }

    // Saving progress callback that aborts after a very short duration.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly DateTime _start = DateTime.Now;
        private const double MaxDurationSeconds = 0.001; // Very short to force cancellation.

        public void Notify(DocumentSavingArgs args)
        {
            double elapsed = (DateTime.Now - _start).TotalSeconds;
            if (elapsed > MaxDurationSeconds)
                throw new OperationCanceledException($"Saving cancelled. EstimatedProgress = {args.EstimatedProgress}");
        }
    }
}
