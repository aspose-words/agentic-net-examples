using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsCancellationDemo
{
    // Callback that aborts saving after a short time interval.
    class CancelSavingCallback : IDocumentSavingCallback
    {
        private readonly DateTime _startTime = DateTime.Now;
        private const double MaxDurationSeconds = 0.01; // Adjust as needed.

        public void Notify(DocumentSavingArgs args)
        {
            if ((DateTime.Now - _startTime).TotalSeconds > MaxDurationSeconds)
                throw new OperationCanceledException(
                    $"Saving canceled. EstimatedProgress = {args.EstimatedProgress}");
        }
    }

    public class Program
    {
        // Configuration setting: turn cancellation support on or off.
        private static readonly bool EnableCancellation = true;

        public static void Main()
        {
            // Prepare a simple document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world! This document demonstrates cancellation support.");

            // Define output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");

            if (EnableCancellation)
            {
                // Attach a progress callback that may cancel the operation.
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
                {
                    ProgressCallback = new CancelSavingCallback()
                };

                try
                {
                    doc.Save(outputPath, saveOptions);
                    Console.WriteLine("Document saved successfully (cancellation not triggered).");
                }
                catch (OperationCanceledException ex)
                {
                    Console.WriteLine($"Saving was canceled: {ex.Message}");
                }
            }
            else
            {
                // Save without cancellation support.
                doc.Save(outputPath);
                Console.WriteLine("Document saved successfully.");

                // Verify that the file exists.
                if (File.Exists(outputPath))
                    Console.WriteLine($"Output file verified at: {outputPath}");
                else
                    throw new FileNotFoundException("The expected output file was not created.", outputPath);
            }
        }
    }
}
