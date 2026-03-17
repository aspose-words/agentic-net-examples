using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace AsposeWordsCancellationDemo
{
    // Loading callback that checks a shared CancellationToken.
    public class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly CancellationToken _cancellationToken;

        public LoadingProgressCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
        }

        public void Notify(DocumentLoadingArgs args)
        {
            // Abort loading if cancellation was requested.
            if (_cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException("Document loading was cancelled.");
        }
    }

    // Saving callback that checks the same CancellationToken.
    public class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _cancellationToken;

        public SavingProgressCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
        }

        public void Notify(DocumentSavingArgs args)
        {
            // Abort saving if cancellation was requested.
            if (_cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException($"Document saving was cancelled at {args.EstimatedProgress}% progress.");
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a CancellationTokenSource that can be used to cancel both operations.
            using var cts = new CancellationTokenSource();

            // Example: cancel after 500 ms (adjust as needed).
            cts.CancelAfter(TimeSpan.FromMilliseconds(500));
            CancellationToken token = cts.Token;

            // Prepare load options with the shared token callback.
            LoadOptions loadOptions = new LoadOptions
            {
                ProgressCallback = new LoadingProgressCallback(token)
            };

            // Load the document using the load options.
            Document doc;
            try
            {
                doc = new Document("Input.docx", loadOptions);
            }
            catch (OperationCanceledException ex)
            {
                Console.WriteLine($"Loading aborted: {ex.Message}");
                return;
            }

            // Prepare save options with the same token callback.
            SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Docx);
            saveOptions.ProgressCallback = new SavingProgressCallback(token);

            // Save the document using the save options.
            try
            {
                doc.Save("Output.docx", saveOptions);
                Console.WriteLine("Document saved successfully.");
            }
            catch (OperationCanceledException ex)
            {
                Console.WriteLine($"Saving aborted: {ex.Message}");
            }
        }
    }
}
