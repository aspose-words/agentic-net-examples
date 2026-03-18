using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace AsposeWordsDebugExample
{
    // Callback that aborts loading after a short duration.
    public class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly DateTime _loadingStartedAt = DateTime.Now;
        private const double MaxDuration = 0.5; // seconds

        public void Notify(DocumentLoadingArgs args)
        {
            double elapsedSeconds = (DateTime.Now - _loadingStartedAt).TotalSeconds;
            if (elapsedSeconds > MaxDuration)
                throw new OperationCanceledException(
                    $"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
        }
    }

    // Callback that aborts saving after a short duration.
    public class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly DateTime _savingStartedAt = DateTime.Now;
        private const double MaxDuration = 0.1; // seconds

        public void Notify(DocumentSavingArgs args)
        {
            double elapsedSeconds = (DateTime.Now - _savingStartedAt).TotalSeconds;
            if (elapsedSeconds > MaxDuration)
                throw new OperationCanceledException(
                    $"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Paths to input and output files (adjust as needed).
            string inputPath = Path.Combine(Environment.CurrentDirectory, "Big document.docx");
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Result.docx");

            // ---------- Document Loading ----------
            var loadOptions = new LoadOptions
            {
                ProgressCallback = new LoadingProgressCallback()
            };

            Document doc = null;
            try
            {
                doc = new Document(inputPath, loadOptions);
                Console.WriteLine("Document loaded successfully.");
            }
            catch (OperationCanceledException ex)
            {
                // Log the exception message and full stack trace for debugging.
                Console.WriteLine("Loading was canceled:");
                Console.WriteLine(ex.Message);
                Console.WriteLine("Stack Trace:");
                Console.WriteLine(ex.StackTrace);
                // Optionally rethrow or handle further.
                return;
            }

            // ---------- Document Saving ----------
            var saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ProgressCallback = new SavingProgressCallback()
            };

            try
            {
                doc.Save(outputPath, saveOptions);
                Console.WriteLine("Document saved successfully.");
            }
            catch (OperationCanceledException ex)
            {
                // Log the exception message and full stack trace for debugging.
                Console.WriteLine("Saving was canceled:");
                Console.WriteLine(ex.Message);
                Console.WriteLine("Stack Trace:");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
