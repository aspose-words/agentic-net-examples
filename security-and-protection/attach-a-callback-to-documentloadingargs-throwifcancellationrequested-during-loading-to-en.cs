using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

namespace AsposeLoadingCallbackDemo
{
    // Callback that aborts loading if the operation takes longer than the allowed duration.
    class LoadingCallback : IDocumentLoadingCallback
    {
        private readonly DateTime _loadingStartedAt;
        private const double MaxDurationSeconds = 0.5; // Cancel after half a second.

        public LoadingCallback()
        {
            _loadingStartedAt = DateTime.Now;
        }

        public void Notify(DocumentLoadingArgs args)
        {
            // Determine how long loading has been running.
            double elapsed = (DateTime.Now - _loadingStartedAt).TotalSeconds;
            if (elapsed > MaxDurationSeconds)
            {
                // Abort loading by throwing an OperationCanceledException.
                // Include progress information for diagnostic purposes.
                throw new OperationCanceledException(
                    $"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare a temporary folder for the sample files.
            string artifactsDir = Path.Combine(Path.GetTempPath(), "AsposeLoadingCallbackDemo");
            Directory.CreateDirectory(artifactsDir);

            // Create a simple document and save it.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world! This document will be loaded with a cancellation callback.");
            string filePath = Path.Combine(artifactsDir, "Sample.docx");
            doc.Save(filePath);

            // Set up load options with the custom progress callback.
            LoadOptions loadOptions = new LoadOptions
            {
                ProgressCallback = new LoadingCallback()
            };

            // Attempt to load the document; expect an OperationCanceledException if cancelled.
            try
            {
                Document loadedDoc = new Document(filePath, loadOptions);
                Console.WriteLine("Document loaded successfully.");
            }
            catch (OperationCanceledException ex)
            {
                Console.WriteLine($"Loading was cancelled: {ex.Message}");
            }

            // Clean up the temporary file.
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }
}
