using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Custom saving progress callback that monitors a CancellationToken.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _cancellationToken;
        private readonly DateTime _startTime;

        public SavingProgressCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
            _startTime = DateTime.Now;
        }

        public void Notify(DocumentSavingArgs args)
        {
            // If cancellation was requested, abort the save operation.
            if (_cancellationToken.IsCancellationRequested)
            {
                throw new OperationCanceledException(
                    $"Save operation canceled at {args.EstimatedProgress:F2}% progress.");
            }

            // Optional: simulate a long operation by checking elapsed time.
            // This example does not delay, but the check above allows early exit.
        }
    }

    public static void Main()
    {
        // Prepare output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ProcessedDocument.docx");

        // Create a sample document with enough content to generate progress.
        Document doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (int i = 0; i < 500; i++)
        {
            builder.Writeln($"This is line {i + 1} of the sample document.");
        }

        // Set up a cancellation token that will be triggered after a short delay.
        using var cts = new CancellationTokenSource();
        Task.Run(async () =>
        {
            await Task.Delay(100); // Cancel after 100 ms.
            cts.Cancel();
        });

        // Configure save options with the progress callback.
        var saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new SavingProgressCallback(cts.Token)
        };

        try
        {
            // Attempt to save the document. The callback may throw OperationCanceledException.
            doc.Save(outputPath, saveOptions);
            Console.WriteLine("Document saved successfully.");
        }
        catch (OperationCanceledException ex)
        {
            // Inform the user that the operation was canceled and provide progress info.
            Console.WriteLine($"Saving was canceled. Details: {ex.Message}");
        }
        finally
        {
            // Clean up the partially saved file if it exists.
            if (File.Exists(outputPath))
            {
                try { File.Delete(outputPath); } catch { /* ignore cleanup errors */ }
            }
        }
    }
}
