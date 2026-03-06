using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

// Callback that monitors a CancellationToken during document loading.
class CancellationTokenLoadingCallback : IDocumentLoadingCallback
{
    private readonly CancellationToken _token;

    public CancellationTokenLoadingCallback(CancellationToken token)
    {
        _token = token;
    }

    // Called periodically by Aspose.Words while loading.
    public void Notify(DocumentLoadingArgs args)
    {
        // If cancellation is requested, abort loading.
        if (_token.IsCancellationRequested)
            throw new OperationCanceledException($"Loading canceled at progress {args.EstimatedProgress}%");
    }
}

// Callback that monitors a CancellationToken during document saving.
class CancellationTokenSavingCallback : IDocumentSavingCallback
{
    private readonly CancellationToken _token;

    public CancellationTokenSavingCallback(CancellationToken token)
    {
        _token = token;
    }

    // Called periodically by Aspose.Words while saving.
    public void Notify(DocumentSavingArgs args)
    {
        // If cancellation is requested, abort saving.
        if (_token.IsCancellationRequested)
            throw new OperationCanceledException($"Saving canceled at progress {args.EstimatedProgress}%");
    }
}

class Program
{
    static void Main()
    {
        // Create a token source that can be cancelled from another thread or timer.
        var cts = new CancellationTokenSource();

        // Example: automatically cancel after 200 ms (replace with real cancellation logic).
        var timer = new System.Timers.Timer(200);
        timer.Elapsed += (s, e) => { cts.Cancel(); timer.Stop(); };
        timer.Start();

        try
        {
            // Load the DOCX file with a progress callback that checks the token.
            var loadOptions = new LoadOptions
            {
                ProgressCallback = new CancellationTokenLoadingCallback(cts.Token)
            };
            Document doc = new Document("Input.docx", loadOptions);

            // Perform any document processing here.
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Processing completed.");

            // Save the document with a progress callback that checks the token.
            var saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                ProgressCallback = new CancellationTokenSavingCallback(cts.Token)
            };
            doc.Save("Output.docx", saveOptions);
        }
        catch (OperationCanceledException ex)
        {
            // Handle the cancellation gracefully.
            Console.WriteLine($"Operation was canceled: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handle other unexpected errors.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
