using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsResourceLeakDemo
{
    // Callback that aborts the save operation after a short delay.
    public class CancelSavingCallback : IDocumentSavingCallback
    {
        private readonly DateTime _startTime = DateTime.Now;
        private const double MaxDurationSeconds = 0.01; // Cancel almost immediately.

        public void Notify(DocumentSavingArgs args)
        {
            double elapsed = (DateTime.Now - _startTime).TotalSeconds;
            if (elapsed > MaxDurationSeconds)
                throw new OperationCanceledException(
                    $"EstimatedProgress = {args.EstimatedProgress}; Canceled after {elapsed:F2}s");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path for the output document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CanceledSave.docx");

            // Ensure any previous file is removed.
            if (File.Exists(outputPath))
                File.Delete(outputPath);

            Document doc = null;
            try
            {
                // Create a new blank document.
                doc = new Document();

                // Add simple content.
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Writeln("Hello Aspose.Words!");

                // Configure save options with a progress callback that will cancel the operation.
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
                {
                    ProgressCallback = new CancelSavingCallback()
                };

                // Attempt to save; the callback will throw OperationCanceledException.
                doc.Save(outputPath, saveOptions);
            }
            catch (OperationCanceledException ex)
            {
                // Handle the cancellation gracefully.
                Console.WriteLine($"Document saving was canceled: {ex.Message}");
            }
            finally
            {
                // Dispose the Document if it implements IDisposable to prevent resource leaks.
                if (doc is IDisposable disposable)
                    disposable.Dispose();
            }

            // Verify that the file was not created due to cancellation.
            bool fileExists = File.Exists(outputPath);
            Console.WriteLine($"Output file exists: {fileExists}");
        }
    }
}
