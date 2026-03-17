using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class DocumentLoader
{
    // Maximum allowed loading duration in seconds before cancellation.
    private const double MaxDurationSeconds = 0.5;

    // Temporary folder used during document loading.
    private readonly string _tempFolder = Path.Combine(Path.GetTempPath(), "AsposeLoadTemp");

    public void LoadDocumentWithCancellation(string filePath)
    {
        // Ensure the temporary folder exists before loading.
        Directory.CreateDirectory(_tempFolder);

        // Configure load options with a progress callback and temporary folder.
        LoadOptions loadOptions = new LoadOptions
        {
            ProgressCallback = new LoadingProgressCallback(),
            TempFolder = _tempFolder
        };

        try
        {
            // Attempt to load the document. The callback may throw OperationCanceledException.
            Document doc = new Document(filePath, loadOptions);

            // Document loaded successfully – further processing can be done here.
            Console.WriteLine("Document loaded successfully.");
        }
        catch (OperationCanceledException ex)
        {
            // Handle the cancellation: log the reason.
            Console.WriteLine($"Document loading was cancelled: {ex.Message}");
        }
        finally
        {
            // Cleanup any temporary files created during the load operation.
            try
            {
                if (Directory.Exists(_tempFolder))
                {
                    Directory.Delete(_tempFolder, true);
                }
            }
            catch (Exception cleanupEx)
            {
                // Log cleanup failures without interrupting the application flow.
                Console.WriteLine($"Failed to delete temporary folder: {cleanupEx.Message}");
            }
        }
    }

    // Implements IDocumentLoadingCallback to monitor loading progress and abort if it exceeds MaxDurationSeconds.
    private class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly DateTime _startTime = DateTime.Now;

        public void Notify(DocumentLoadingArgs args)
        {
            double elapsedSeconds = (DateTime.Now - _startTime).TotalSeconds;
            if (elapsedSeconds > MaxDurationSeconds)
                throw new OperationCanceledException($"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
        }
    }
}

public static class Program
{
    // Entry point required by the C# compiler.
    public static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: dotnet run <path-to-document>");
            return;
        }

        string filePath = args[0];
        var loader = new DocumentLoader();
        loader.LoadDocumentWithCancellation(filePath);
    }
}
