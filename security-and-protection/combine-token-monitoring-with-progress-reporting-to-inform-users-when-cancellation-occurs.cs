using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Custom progress callback that also monitors a cancellation token.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _cancellationToken;

        public SavingProgressCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
        }

        public void Notify(DocumentSavingArgs args)
        {
            // Report progress to the console.
            Console.WriteLine($"Saving progress: {args.EstimatedProgress:F2}%");

            // If cancellation was requested, abort the save operation.
            if (_cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException(
                    $"Save operation cancelled at {args.EstimatedProgress:F2}% progress.");
        }
    }

    public static void Main()
    {
        // Prepare a large document to make the save operation take noticeable time.
        Document doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (int i = 0; i < 2000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SavedDocument.docx");

        // Ensure any previous file is removed.
        if (File.Exists(outputPath))
            File.Delete(outputPath);

        // Set up a cancellation token that will trigger after a short delay.
        using var cts = new CancellationTokenSource();
        // Cancel after 200 milliseconds.
        cts.CancelAfter(200);

        // Configure save options with the custom progress callback.
        var saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new SavingProgressCallback(cts.Token)
        };

        try
        {
            // Perform the save operation. This will invoke the progress callback repeatedly.
            doc.Save(outputPath, saveOptions);
            Console.WriteLine("Document saved successfully.");
        }
        catch (OperationCanceledException ex)
        {
            // Inform the user that the operation was cancelled.
            Console.WriteLine($"Operation was cancelled: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Any other unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
