using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

namespace AsposeLoadingCancellationDemo
{
    // Custom callback that aborts loading when the elapsed time exceeds MaxDuration.
    public class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly DateTime _loadingStartedAt;

        // Set MaxDuration to 0 seconds to cancel immediately on the first callback.
        private const double MaxDuration = 0.0;

        public LoadingProgressCallback()
        {
            _loadingStartedAt = DateTime.Now;
        }

        public void Notify(DocumentLoadingArgs args)
        {
            double elapsedSeconds = (DateTime.Now - _loadingStartedAt).TotalSeconds;
            if (elapsedSeconds > MaxDuration)
                throw new OperationCanceledException(
                    $"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare a temporary folder for the sample document.
            string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeLoadingDemo");
            Directory.CreateDirectory(tempFolder);

            string docPath = Path.Combine(tempFolder, "Sample.docx");

            // Create a simple document and save it.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello Aspose.Words!");
            doc.Save(docPath);

            // Configure LoadOptions with the custom progress callback.
            LoadOptions loadOptions = new LoadOptions
            {
                ProgressCallback = new LoadingProgressCallback()
            };

            try
            {
                // Attempt to load the document; the callback will cancel the operation.
                Document loadedDoc = new Document(docPath, loadOptions);
                Console.WriteLine("Document loaded successfully (unexpected).");
            }
            catch (OperationCanceledException ex)
            {
                Console.WriteLine($"Loading was canceled: {ex.Message}");
            }
            finally
            {
                // Clean up temporary files.
                if (File.Exists(docPath))
                    File.Delete(docPath);
                if (Directory.Exists(tempFolder))
                    Directory.Delete(tempFolder, true);
            }
        }
    }
}
