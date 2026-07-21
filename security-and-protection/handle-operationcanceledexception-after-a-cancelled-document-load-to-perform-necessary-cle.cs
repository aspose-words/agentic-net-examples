using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

namespace AsposeWordsCancellationExample
{
    // Callback that cancels document loading after a short duration.
    public class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly DateTime _loadingStartedAt = DateTime.Now;
        private const double MaxDurationSeconds = 0.1; // Cancel after 0.1 seconds.

        public void Notify(DocumentLoadingArgs args)
        {
            double elapsedSeconds = (DateTime.Now - _loadingStartedAt).TotalSeconds;
            if (elapsedSeconds > MaxDurationSeconds)
                throw new OperationCanceledException(
                    $"Loading canceled. EstimatedProgress = {args.EstimatedProgress}; Elapsed = {elapsedSeconds}s");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare a temporary folder for the sample files.
            string tempDir = Path.Combine(Path.GetTempPath(), "AsposeWordsCancellation");
            Directory.CreateDirectory(tempDir);

            // Path of the sample document.
            string docPath = Path.Combine(tempDir, "Sample.docx");

            // Create a simple document and save it.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document used to demonstrate loading cancellation.");
            doc.Save(docPath);

            // Set up load options with the progress callback that will cancel loading.
            LoadOptions loadOptions = new LoadOptions
            {
                ProgressCallback = new LoadingProgressCallback()
            };

            try
            {
                // Attempt to load the document. The callback will throw OperationCanceledException.
                Document loadedDoc = new Document(docPath, loadOptions);
                // If loading succeeds (unlikely), write a message.
                Console.WriteLine("Document loaded successfully (unexpected).");
            }
            catch (OperationCanceledException ex)
            {
                // Handle the cancellation and perform necessary cleanup.
                Console.WriteLine($"Loading was canceled: {ex.Message}");
            }
            finally
            {
                // Cleanup: delete the temporary files and folder.
                try
                {
                    if (File.Exists(docPath))
                        File.Delete(docPath);
                    if (Directory.Exists(tempDir))
                        Directory.Delete(tempDir, true);
                }
                catch
                {
                    // Suppress any cleanup exceptions.
                }
            }
        }
    }
}
