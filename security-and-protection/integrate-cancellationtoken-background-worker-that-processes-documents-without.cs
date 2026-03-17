using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class DocumentProcessor
{
    // Asynchronously loads, processes, and saves a document without blocking the UI.
    // The operation can be cancelled via the provided CancellationToken.
    public static Task ProcessDocumentAsync(string inputFile, string outputFile, CancellationToken cancellationToken)
    {
        // Run the whole workflow on a background thread.
        return Task.Run(() =>
        {
            // ---------- Loading ----------
            var loadOptions = new LoadOptions
            {
                ProgressCallback = new LoadingProgressCallback(cancellationToken)
            };

            Document doc = new Document(inputFile, loadOptions);

            // ---------- (Optional) Document processing ----------
            var builder = new DocumentBuilder(doc);
            builder.Writeln($"Processed at {DateTime.Now}");

            // ---------- Saving ----------
            var saveOptions = new HtmlSaveOptions
            {
                ProgressCallback = new SavingProgressCallback(cancellationToken)
            };

            doc.Save(outputFile, saveOptions);
        }, cancellationToken);
    }

    private class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly CancellationToken _cancellationToken;
        private readonly DateTime _startTime;

        public LoadingProgressCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
            _startTime = DateTime.Now;
        }

        public void Notify(DocumentLoadingArgs args)
        {
            if (_cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException($"Loading cancelled at {DateTime.Now}, progress={args.EstimatedProgress}%.");

            const double maxDurationSeconds = 30.0;
            if ((DateTime.Now - _startTime).TotalSeconds > maxDurationSeconds)
                throw new OperationCanceledException($"Loading timeout after {maxDurationSeconds}s, progress={args.EstimatedProgress}%.");
        }
    }

    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _cancellationToken;
        private readonly DateTime _startTime;

        public SavingProgressCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
            _startTime = DateTime.Now;
        }

        public void Notify(DocumentSavingArgs args)
        {
            if (_cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException($"Saving cancelled at {DateTime.Now}, progress={args.EstimatedProgress}%.");

            const double maxDurationSeconds = 30.0;
            if ((DateTime.Now - _startTime).TotalSeconds > maxDurationSeconds)
                throw new OperationCanceledException($"Saving timeout after {maxDurationSeconds}s, progress={args.EstimatedProgress}%.");
        }
    }
}

public static class Program
{
    // Entry point required for a console application. It demonstrates how to start the background
    // processing and how to cancel it (e.g., after a user presses a key).
    public static async Task Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: <inputFile> <outputFile>");
            return;
        }

        string inputFile = args[0];
        string outputFile = args[1];

        using var cts = new CancellationTokenSource();
        Console.CancelKeyPress += (s, e) =>
        {
            e.Cancel = true; // prevent the process from terminating immediately
            cts.Cancel();
            Console.WriteLine("Cancellation requested...");
        };

        try
        {
            await DocumentProcessor.ProcessDocumentAsync(inputFile, outputFile, cts.Token);
            Console.WriteLine("Document processed successfully.");
        }
        catch (OperationCanceledException ex)
        {
            Console.WriteLine($"Operation cancelled: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
