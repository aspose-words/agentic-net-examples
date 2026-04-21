using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsResourceLeakDemo
{
    // Callback that aborts the saving process by throwing OperationCanceledException.
    public class SavingProgressCallback : IDocumentSavingCallback
    {
        public void Notify(DocumentSavingArgs args)
        {
            // Immediately cancel the operation.
            throw new OperationCanceledException($"Saving canceled at progress {args.EstimatedProgress}%.");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare output path.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "ProtectedDocument.docx");

            // Ensure any previous file is removed.
            if (File.Exists(outputPath))
                File.Delete(outputPath);

            // Create the Document without a using block (Document does not implement IDisposable).
            Document doc = new Document();
            try
            {
                // Add simple content.
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Writeln("This document will attempt to save, but the operation will be canceled.");

                // Set up save options with the progress callback that throws.
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
                {
                    ProgressCallback = new SavingProgressCallback()
                };

                // Attempt to save – the callback will cancel the operation.
                doc.Save(outputPath, saveOptions);
            }
            catch (OperationCanceledException ex)
            {
                // Handle the cancellation.
                Console.WriteLine($"Save operation was canceled: {ex.Message}");
            }
            finally
            {
                // Document does not implement IDisposable, so no explicit Dispose call is required.
                // Setting the reference to null allows the garbage collector to reclaim the object.
                doc = null;
            }

            // Verify that the file was not created due to cancellation.
            if (File.Exists(outputPath))
                Console.WriteLine("Unexpected: the file was created.");
            else
                Console.WriteLine("File was not created as expected after cancellation.");
        }
    }
}
