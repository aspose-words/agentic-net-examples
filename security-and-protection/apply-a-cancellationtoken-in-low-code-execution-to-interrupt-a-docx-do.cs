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
        // Create a cancellation token source that will cancel after 500 ms.
        using var cts = new CancellationTokenSource(500);

        // Attach a loading callback that checks the token.
        var loadOptions = new LoadOptions
        {
            ProgressCallback = new LoadingCancellationCallback(cts.Token)
        };

        try
        {
            // Load the document with the custom loading callback.
            Document doc = new Document("Big document.docx", loadOptions);

            // Attach a saving callback that also checks the token.
            var saveOptions = new PdfSaveOptions
            {
                ProgressCallback = new SavingCancellationCallback(cts.Token)
            };

            // Save the document; the callback may abort the operation.
            doc.Save("output.pdf", saveOptions);
        }
        catch (OperationCanceledException ex)
        {
            // Handle the cancellation.
            Console.WriteLine($"Processing was canceled: {ex.Message}");
        }
    }
}

// Implements IDocumentLoadingCallback and aborts loading when the token is cancelled.
class LoadingCancellationCallback : IDocumentLoadingCallback
{
    private readonly CancellationToken _token;

    public LoadingCancellationCallback(CancellationToken token)
    {
        _token = token;
    }

    public void Notify(DocumentLoadingArgs args)
    {
        if (_token.IsCancellationRequested)
            throw new OperationCanceledException(
                $"Loading canceled at {DateTime.Now}; EstimatedProgress = {args.EstimatedProgress}");
    }
}

// Implements IDocumentSavingCallback and aborts saving when the token is cancelled.
class SavingCancellationCallback : IDocumentSavingCallback
{
    private readonly CancellationToken _token;

    public SavingCancellationCallback(CancellationToken token)
    {
        _token = token;
    }

    public void Notify(DocumentSavingArgs args)
    {
        if (_token.IsCancellationRequested)
            throw new OperationCanceledException(
                $"Saving canceled at {DateTime.Now}; EstimatedProgress = {args.EstimatedProgress}");
    }
}
