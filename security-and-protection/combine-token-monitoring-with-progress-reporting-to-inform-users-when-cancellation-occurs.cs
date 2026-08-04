using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Add many paragraphs to make the save operation take noticeable time.
        for (int i = 0; i < 2000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        // Set up a cancellation token that will be triggered after a short delay.
        using var cts = new CancellationTokenSource();
        // Cancel after 200 milliseconds.
        Task.Delay(200).ContinueWith(_ => cts.Cancel());

        // Configure save options with a progress callback that monitors the token.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new SavingProgressCallback(cts.Token)
        };

        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ProcessedDocument.docx");

        try
        {
            // Attempt to save the document. The callback may throw if cancellation occurs.
            doc.Save(outputPath, saveOptions);
            Console.WriteLine("Document saved successfully.");
        }
        catch (OperationCanceledException ex)
        {
            // Inform the user that the operation was canceled and provide progress info.
            Console.WriteLine($"Saving was canceled. Details: {ex.Message}");
        }

        // Verify whether the output file exists (it may be incomplete if canceled).
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Output file exists at: {outputPath}");
        }
        else
        {
            Console.WriteLine("Output file was not created.");
        }
    }

    // Implements IDocumentSavingCallback to receive progress notifications during saving.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _cancellationToken;

        public SavingProgressCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
        }

        public void Notify(DocumentSavingArgs args)
        {
            // If cancellation has been requested, abort the save operation.
            if (_cancellationToken.IsCancellationRequested)
            {
                throw new OperationCanceledException(
                    $"EstimatedProgress = {args.EstimatedProgress}; Save operation was canceled.");
            }
        }
    }
}
