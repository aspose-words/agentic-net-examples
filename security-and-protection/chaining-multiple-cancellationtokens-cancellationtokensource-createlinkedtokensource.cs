using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    // Callback for document loading that respects a linked cancellation token.
    private class LinkedLoadingCallback : IDocumentLoadingCallback
    {
        private readonly CancellationToken _token;
        private readonly DateTime _startedAt = DateTime.Now;

        public LinkedLoadingCallback(CancellationToken token) => _token = token;

        public void Notify(DocumentLoadingArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException($"Loading cancelled at {DateTime.Now}");

            if ((DateTime.Now - _startedAt).TotalSeconds > 5)
                throw new OperationCanceledException($"Loading timeout after 5 seconds. EstimatedProgress = {args.EstimatedProgress}");
        }
    }

    // Callback for document saving that respects a linked cancellation token.
    private class LinkedSavingCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _token;
        private readonly DateTime _startedAt = DateTime.Now;

        public LinkedSavingCallback(CancellationToken token) => _token = token;

        public void Notify(DocumentSavingArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException($"Saving cancelled at {DateTime.Now}");

            if ((DateTime.Now - _startedAt).TotalSeconds > 5)
                throw new OperationCanceledException($"Saving timeout after 5 seconds. EstimatedProgress = {args.EstimatedProgress}");
        }
    }

    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create two independent CancellationTokenSource instances.
        // -----------------------------------------------------------------
        var ctsTimeout = new CancellationTokenSource();
        var ctsUser = new CancellationTokenSource();

        // Simulate a timeout cancellation after 200 ms.
        var timer = new Timer(_ => ctsTimeout.Cancel(), null, 200, Timeout.Infinite);

        // Simulate a user‑initiated cancellation after 500 ms.
        var userTimer = new Timer(_ => ctsUser.Cancel(), null, 500, Timeout.Infinite);

        // -----------------------------------------------------------------
        // 2. Link the tokens so that cancellation of any source aborts the workflow.
        // -----------------------------------------------------------------
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(ctsTimeout.Token, ctsUser.Token);
        CancellationToken linkedToken = linkedCts.Token;

        // -----------------------------------------------------------------
        // 3. Load a document using a loading callback that checks the linked token.
        // -----------------------------------------------------------------
        Document doc;
        try
        {
            var loadCallback = new LinkedLoadingCallback(linkedToken);
            var loadOptions = new LoadOptions { ProgressCallback = loadCallback };

            // Create a simple document in memory to avoid external file dependencies.
            var sampleDoc = new Document();
            var builder = new Aspose.Words.DocumentBuilder(sampleDoc);
            builder.Writeln("Sample text for cancellation demo.");

            using var ms = new MemoryStream();
            sampleDoc.Save(ms, SaveFormat.Docx);
            ms.Position = 0;

            doc = new Document(ms, loadOptions);
        }
        catch (OperationCanceledException ex)
        {
            Console.WriteLine($"Document loading was aborted: {ex.Message}");
            return;
        }

        // -----------------------------------------------------------------
        // 4. Save the document using a saving callback that also checks the linked token.
        // -----------------------------------------------------------------
        try
        {
            var saveCallback = new LinkedSavingCallback(linkedToken);
            var saveOptions = new HtmlSaveOptions(SaveFormat.Html) { ProgressCallback = saveCallback };

            doc.Save("Output.html", saveOptions);
            Console.WriteLine("Document saved successfully to Output.html");
        }
        catch (OperationCanceledException ex)
        {
            Console.WriteLine($"Document saving was aborted: {ex.Message}");
        }
        finally
        {
            timer.Dispose();
            userTimer.Dispose();
        }
    }
}
