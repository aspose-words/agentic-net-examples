using System;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsCancellationDemo
{
    // Callback that aborts saving when the supplied CancellationToken is signaled.
    public class TokenSavingCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _cancellationToken;

        public TokenSavingCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
        }

        // This method is called periodically during document saving.
        public void Notify(DocumentSavingArgs args)
        {
            if (_cancellationToken.IsCancellationRequested)
                // Throwing OperationCanceledException aborts the save operation.
                throw new OperationCanceledException(
                    $"Saving canceled. EstimatedProgress = {args.EstimatedProgress}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a cancellation token that will be triggered after 200 milliseconds.
            using var cts = new CancellationTokenSource();
            cts.CancelAfter(200);

            // Load the source DOCX document.
            Document doc = new Document("InputDocument.docx");

            // Configure save options and attach the cancellation callback.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                ProgressCallback = new TokenSavingCallback(cts.Token)
            };

            try
            {
                // Attempt to save the document; may be canceled by the token.
                doc.Save("OutputDocument.docx", saveOptions);
                Console.WriteLine("Document saved successfully.");
            }
            catch (OperationCanceledException ex)
            {
                // Handle the cancellation gracefully.
                Console.WriteLine($"Save operation was canceled: {ex.Message}");
            }
        }
    }
}
