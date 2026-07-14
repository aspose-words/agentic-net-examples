using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsResourceLeakDemo
{
    // Callback that aborts the save operation after a short duration.
    public class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly DateTime _savingStartedAt = DateTime.Now;
        private const double MaxDurationSeconds = 0.01; // Cancel quickly.

        public void Notify(DocumentSavingArgs args)
        {
            double elapsed = (DateTime.Now - _savingStartedAt).TotalSeconds;
            if (elapsed > MaxDurationSeconds)
                throw new OperationCanceledException($"EstimatedProgress = {args.EstimatedProgress}");
        }
    }

    public class Program
    {
        public static void Main()
        {
            const string outputPath = "CanceledDocument.docx";
            Document doc = null;

            try
            {
                // Create a simple document.
                doc = new Document();
                var builder = new DocumentBuilder(doc);
                builder.Writeln("Hello world!");

                // Set up save options with the progress callback that will cancel the operation.
                var saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
                {
                    ProgressCallback = new SavingProgressCallback()
                };

                // Attempt to save; this will be canceled and throw OperationCanceledException.
                doc.Save(outputPath, saveOptions);
            }
            catch (OperationCanceledException ex)
            {
                Console.WriteLine($"Save operation was canceled: {ex.Message}");
            }
            finally
            {
                // Document does not implement IDisposable, so no explicit disposal is required.
                // Setting the reference to null allows the garbage collector to reclaim it.
                doc = null;
            }

            // Simple verification that the file was not created due to cancellation.
            if (File.Exists(outputPath))
                Console.WriteLine("File was created (partial save).");
            else
                Console.WriteLine("File was not created because the save was canceled.");
        }
    }
}
