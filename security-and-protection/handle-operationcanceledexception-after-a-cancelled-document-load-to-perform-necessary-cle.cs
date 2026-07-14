using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

namespace AsposeWordsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a temporary folder and sample document.
            string folder = Path.Combine(Path.GetTempPath(), "AsposeExample");
            Directory.CreateDirectory(folder);
            string filePath = Path.Combine(folder, "sample.docx");

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document.");
            doc.Save(filePath);

            // Set up load options with a progress callback that will cancel loading.
            LoadOptions loadOptions = new LoadOptions
            {
                ProgressCallback = new LoadingProgressCallback()
            };

            Document loadedDoc = null;
            try
            {
                // Attempt to load the document; the callback will throw OperationCanceledException.
                loadedDoc = new Document(filePath, loadOptions);
            }
            catch (OperationCanceledException ex)
            {
                // Handle the cancellation and perform cleanup.
                Console.WriteLine($"Loading cancelled: {ex.Message}");

                // Cleanup: delete the source file if it exists.
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }

            // Remove the temporary folder.
            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, true);
            }
        }

        // Callback that cancels loading after a short duration.
        private class LoadingProgressCallback : IDocumentLoadingCallback
        {
            private readonly DateTime _start = DateTime.Now;
            private const double MaxDurationSeconds = 0.1; // Cancel quickly.

            public void Notify(DocumentLoadingArgs args)
            {
                double elapsed = (DateTime.Now - _start).TotalSeconds;
                if (elapsed > MaxDurationSeconds)
                {
                    throw new OperationCanceledException(
                        $"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
                }
            }
        }
    }
}
