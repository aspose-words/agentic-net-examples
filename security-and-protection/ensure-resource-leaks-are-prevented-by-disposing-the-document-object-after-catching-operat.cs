using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsResourceLeakDemo
{
    // Callback that cancels the save operation after a short duration.
    class CancelSavingCallback : IDocumentSavingCallback
    {
        private readonly DateTime _startTime = DateTime.Now;
        private const double MaxDurationSeconds = 0.01; // Cancel quickly.

        public void Notify(DocumentSavingArgs args)
        {
            if ((DateTime.Now - _startTime).TotalSeconds > MaxDurationSeconds)
                throw new OperationCanceledException($"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
        }
    }

    public class Program
    {
        public static void Main()
        {
            const string outputPath = "DemoDocument.docx";

            // Create a simple document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello Aspose.Words!");

            // Prepare save options with a progress callback that will abort the operation.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                ProgressCallback = new CancelSavingCallback()
            };

            try
            {
                // Attempt to save; the callback will throw OperationCanceledException.
                doc.Save(outputPath, saveOptions);
            }
            catch (OperationCanceledException ex)
            {
                Console.WriteLine($"Save operation was canceled: {ex.Message}");
            }

            // Verify that the file was not created due to cancellation.
            if (File.Exists(outputPath))
                Console.WriteLine("File was created despite cancellation.");
            else
                Console.WriteLine("File was not created, as expected.");
        }
    }
}
