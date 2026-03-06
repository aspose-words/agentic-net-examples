using System;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

class TokenCancellationCallback : IDocumentSavingCallback
{
    private readonly CancellationToken _token;

    public TokenCancellationCallback(CancellationToken token) => _token = token;

    // Called during document saving; abort if cancellation is requested.
    public void Notify(DocumentSavingArgs args)
    {
        if (_token.IsCancellationRequested)
            throw new OperationCanceledException($"EstimatedProgress = {args.EstimatedProgress}");
    }
}

class Program
{
    static void Main()
    {
        // Cancel the operation after 100 milliseconds.
        using var cts = new CancellationTokenSource(100);

        // Load an existing DOCX document (uses Document constructor rule).
        Document doc = new Document("Input.docx");

        // Configure save options and attach the cancellation callback (uses SaveOptions rule).
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new TokenCancellationCallback(cts.Token)
        };

        try
        {
            // Save the document with the configured options (uses Document.Save rule).
            doc.Save("Output.docx", saveOptions);
        }
        catch (OperationCanceledException ex)
        {
            Console.WriteLine($"Document saving was canceled: {ex.Message}");
        }
    }
}
