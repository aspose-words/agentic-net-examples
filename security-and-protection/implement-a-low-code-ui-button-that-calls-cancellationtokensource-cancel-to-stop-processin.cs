using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Path for the output document.
        const string outputPath = "output.docx";

        // Clean up any previous run.
        if (File.Exists(outputPath))
            File.Delete(outputPath);

        // Create a sample document with enough content to make saving take noticeable time.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a test document.");
        for (int i = 0; i < 1000; i++)
            builder.Writeln($"Line {i}");

        // Cancellation token source that will be triggered by the simulated UI button.
        CancellationTokenSource cts = new CancellationTokenSource();

        // Simulate a low‑code UI button press that cancels the operation after a short delay.
        Task.Run(async () =>
        {
            await Task.Delay(10); // Delay represents the time before the user clicks the button.
            cts.Cancel();         // Button click invokes Cancel().
        });

        // Configure save options with a progress callback that observes the cancellation token.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new SavingProgressCallback(cts.Token)
        };

        try
        {
            // Attempt to save the document; the operation will be aborted when the token is cancelled.
            doc.Save(outputPath, saveOptions);
        }
        catch (OperationCanceledException)
        {
            // Expected outcome when the cancellation token is triggered.
        }

        // If a partially written file exists, remove it to leave a clean state.
        if (File.Exists(outputPath))
            File.Delete(outputPath);
    }

    // Progress callback that throws an exception when cancellation is requested.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _token;

        public SavingProgressCallback(CancellationToken token) => _token = token;

        public void Notify(DocumentSavingArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException($"Cancelled at progress {args.EstimatedProgress}");
        }
    }
}
