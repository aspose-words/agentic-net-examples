using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the input DOCX file.
        string inputPath = @"C:\Docs\LargeDocument.docx";

        // Create a CancellationTokenSource that will request cancellation after 500 ms.
        using var cts = new CancellationTokenSource();
        cts.CancelAfter(TimeSpan.FromMilliseconds(500));

        // Configure load options with a custom progress callback that checks the token.
        LoadOptions loadOptions = new LoadOptions
        {
            ProgressCallback = new LoadingProgressCallback(cts.Token)
        };

        try
        {
            // Load the document; the callback may throw OperationCanceledException.
            Document doc = new Document(inputPath, loadOptions);

            // Perform any additional processing here (omitted for brevity).

            // Example of saving with a similar cancellation check.
            string outputPath = @"C:\Docs\Processed.docx";
            // Use generic SaveOptions for DOCX format; set the progress callback.
            SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Docx);
            saveOptions.ProgressCallback = new SavingProgressCallback(cts.Token);
            doc.Save(outputPath, saveOptions);
        }
        catch (OperationCanceledException ex)
        {
            Console.WriteLine($"Operation was cancelled: {ex.Message}");
        }
    }

    // Loading callback that aborts when the cancellation token is signaled.
    private class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly CancellationToken _token;
        private readonly DateTime _startedAt;

        public LoadingProgressCallback(CancellationToken token)
        {
            _token = token;
            _startedAt = DateTime.Now;
        }

        public void Notify(DocumentLoadingArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException(
                    $"Loading cancelled at {DateTime.Now}, elapsed {(DateTime.Now - _startedAt).TotalSeconds:F2}s.");
        }
    }

    // Saving callback that aborts when the cancellation token is signaled.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _token;
        private readonly DateTime _startedAt;

        public SavingProgressCallback(CancellationToken token)
        {
            _token = token;
            _startedAt = DateTime.Now;
        }

        public void Notify(DocumentSavingArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException(
                    $"Saving cancelled at {DateTime.Now}, elapsed {(DateTime.Now - _startedAt).TotalSeconds:F2}s.");
        }
    }
}
