using System;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

// Callback that aborts document loading when the supplied token is cancelled.
class LoadingCancellationCallback : IDocumentLoadingCallback
{
    private readonly CancellationToken _token;

    public LoadingCancellationCallback(CancellationToken token) => _token = token;

    public void Notify(DocumentLoadingArgs args)
    {
        if (_token.IsCancellationRequested)
            throw new OperationCanceledException($"Loading canceled at {DateTime.Now}");
    }
}

// Callback that aborts document saving when the supplied token is cancelled.
class SavingCancellationCallback : IDocumentSavingCallback
{
    private readonly CancellationToken _token;

    public SavingCancellationCallback(CancellationToken token) => _token = token;

    public void Notify(DocumentSavingArgs args)
    {
        if (_token.IsCancellationRequested)
            throw new OperationCanceledException($"Saving canceled at {DateTime.Now}");
    }
}

class Program
{
    static void Main()
    {
        // Timeout duration in milliseconds (e.g., 5 seconds).
        const int timeoutMs = 5000;

        // Create a CancellationTokenSource that will cancel automatically after the timeout.
        using var cts = new CancellationTokenSource();
        cts.CancelAfter(timeoutMs);

        // ---------- Load document with timeout ----------
        var loadOptions = new LoadOptions
        {
            ProgressCallback = new LoadingCancellationCallback(cts.Token)
        };

        Document doc;
        try
        {
            // Replace with the actual path to your source document.
            doc = new Document("Big document.docx", loadOptions);
        }
        catch (OperationCanceledException ex)
        {
            Console.WriteLine($"Document loading timed out: {ex.Message}");
            return;
        }

        // ---------- Save document with the same timeout ----------
        var saveOptions = new PdfSaveOptions
        {
            ProgressCallback = new SavingCancellationCallback(cts.Token)
        };

        try
        {
            // Replace with the desired output path.
            doc.Save("Output.pdf", saveOptions);
        }
        catch (OperationCanceledException ex)
        {
            Console.WriteLine($"Document saving timed out: {ex.Message}");
        }
    }
}
