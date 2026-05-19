using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // Output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CanceledSave.docx");

        // Create a blank document and add many paragraphs to make saving take noticeable time.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 2000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        // Set up a cancellation token that will be triggered after a short delay.
        using var cts = new CancellationTokenSource();
        Task.Delay(100).ContinueWith(_ => cts.Cancel());

        // Configure save options and attach a progress callback that observes the token.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new CancellationSavingCallback(cts.Token)
        };

        try
        {
            // Save the document. The callback will throw OperationCanceledException if the token is cancelled.
            doc.Save(outputPath, saveOptions);
            Console.WriteLine("Document saved successfully (cancellation did not occur).");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Document saving was canceled via the cancellation token.");
        }
        finally
        {
            // Remove any partially written file.
            if (File.Exists(outputPath))
            {
                try { File.Delete(outputPath); } catch { /* ignore cleanup errors */ }
            }
        }
    }

    // Implements IDocumentSavingCallback to monitor progress and abort when cancellation is requested.
    private class CancellationSavingCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _token;

        public CancellationSavingCallback(CancellationToken token)
        {
            _token = token;
        }

        public void Notify(DocumentSavingArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException($"Saving canceled at {args.EstimatedProgress}% progress.");
        }
    }
}
