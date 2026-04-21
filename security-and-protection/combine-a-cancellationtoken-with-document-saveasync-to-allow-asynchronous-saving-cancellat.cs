using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static async Task Main(string[] args)
    {
        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document will be saved with cancellation support.");

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AsyncCanceled.docx");

        // Remove any existing file.
        if (File.Exists(outputPath))
            File.Delete(outputPath);

        // Set up save options.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);

        // Create a CancellationTokenSource that cancels after a short delay.
        using CancellationTokenSource cts = new CancellationTokenSource();
        cts.CancelAfter(TimeSpan.FromMilliseconds(100));

        // Attach a progress callback that checks the token and aborts saving.
        saveOptions.ProgressCallback = new CancellationSavingCallback(cts.Token);

        try
        {
            // Save the document. The callback will throw if the token is cancelled.
            await Task.Run(() => doc.Save(outputPath, saveOptions), cts.Token);
        }
        catch (OperationCanceledException)
        {
            // Expected when the token is cancelled.
        }

        // Verify that the file was not created due to cancellation.
        if (File.Exists(outputPath))
            File.Delete(outputPath);
    }

    // Callback that aborts saving when the supplied token is cancelled.
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
                throw new OperationCanceledException($"Saving cancelled at progress {args.EstimatedProgress}%.");
        }
    }
}
