using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsCancellationDemo
{
    // Callback that aborts the save operation after a short time.
    public class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly DateTime _startTime;
        private const double MaxDurationSeconds = 0.05; // Cancel quickly for the demo.

        public SavingProgressCallback()
        {
            _startTime = DateTime.Now;
        }

        public void Notify(DocumentSavingArgs args)
        {
            // If the operation has run longer than the allowed duration, abort it.
            if ((DateTime.Now - _startTime).TotalSeconds > MaxDurationSeconds)
                throw new OperationCanceledException(
                    $"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a sample document with many paragraphs to ensure the save takes noticeable time.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            for (int i = 0; i < 2000; i++)
                builder.Writeln($"Paragraph {i + 1}");

            // Prepare save options with the progress callback that will cancel the operation.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                ProgressCallback = new SavingProgressCallback()
            };

            string outputPath = Path.Combine(Environment.CurrentDirectory, "CanceledDocument.docx");

            try
            {
                // Attempt to save; the callback should abort the process.
                doc.Save(outputPath, saveOptions);
                Console.WriteLine("Document saved successfully (unexpected).");
            }
            catch (OperationCanceledException ex)
            {
                // Expected path: the save operation was cancelled.
                Console.WriteLine($"Save operation was cancelled as expected: {ex.Message}");
            }

            // Verify that the output file does not exist because the operation was aborted.
            if (File.Exists(outputPath))
                Console.WriteLine("Warning: output file was created despite cancellation.");
            else
                Console.WriteLine("No output file was created, confirming cancellation.");
        }
    }
}
