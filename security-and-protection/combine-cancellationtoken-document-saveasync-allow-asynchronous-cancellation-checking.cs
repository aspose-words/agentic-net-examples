using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsCancellationDemo
{
    // Callback that aborts saving when the supplied CancellationToken is cancelled.
    public class CancellationSavingCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _cancellationToken;

        public CancellationSavingCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
        }

        // This method is called periodically during the save operation.
        // If cancellation is requested we throw OperationCanceledException,
        // which Aspose.Words will propagate to the caller.
        public void Notify(DocumentSavingArgs args)
        {
            if (_cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException(
                    $"Saving cancelled at estimated progress {args.EstimatedProgress}%.");
        }
    }

    public static class DocumentExtensions
    {
        // Asynchronous save that respects a CancellationToken.
        public static async Task SaveAsync(this Document document,
                                           string filePath,
                                           SaveOptions saveOptions,
                                           CancellationToken cancellationToken = default)
        {
            // Attach the cancellation callback to the save options.
            saveOptions.ProgressCallback = new CancellationSavingCallback(cancellationToken);

            // Perform the synchronous Save on a background thread.
            await Task.Run(() => document.Save(filePath, saveOptions), cancellationToken)
                      .ConfigureAwait(false);
        }
    }

    class Program
    {
        static async Task Main()
        {
            // Example usage:
            var cancellationSource = new CancellationTokenSource();

            // Load a document (using the standard load rule).
            Document doc = new Document("Input.docx");

            // Prepare save options (any derived SaveOptions can be used).
            var saveOptions = new PdfSaveOptions(); // for example, saving as PDF

            // Start the async save operation.
            Task saveTask = doc.SaveAsync("Output.pdf", saveOptions, cancellationSource.Token);

            // Simulate a condition that triggers cancellation after a short delay.
            await Task.Delay(100);
            cancellationSource.Cancel();

            try
            {
                await saveTask;
                Console.WriteLine("Document saved successfully.");
            }
            catch (OperationCanceledException ex)
            {
                Console.WriteLine($"Save operation was cancelled: {ex.Message}");
            }
        }
    }
}
