using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple document and save it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");
        string filePath = Path.Combine(outputDir, "Sample.docx");
        doc.Save(filePath);

        // Set up load options with a progress callback that forces cancellation.
        LoadOptions loadOptions = new LoadOptions
        {
            ProgressCallback = new LoadingProgressCallback()
        };

        try
        {
            // Attempt to load the document; the callback will trigger cancellation.
            Document loadedDoc = new Document(filePath, loadOptions);
            // If loading succeeds (unlikely), output the document text.
            Console.WriteLine("Document loaded successfully:");
            Console.WriteLine(loadedDoc.GetText());
        }
        catch (OperationCanceledException ex)
        {
            // Expected path when the callback requests cancellation.
            Console.WriteLine("Loading was cancelled by the progress callback:");
            Console.WriteLine(ex.Message);
        }
    }

    // Callback implementation that aborts loading by throwing OperationCanceledException.
    private class LoadingProgressCallback : IDocumentLoadingCallback
    {
        public void Notify(DocumentLoadingArgs args)
        {
            // Throw an exception to abort loading. Include progress info for diagnostics.
            throw new OperationCanceledException(
                $"Loading cancelled. EstimatedProgress = {args.EstimatedProgress}");
        }
    }
}
