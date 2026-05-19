using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Saving;

namespace AsposeWordsCancellationDemo
{
    // Extension method that adds cancellation support to Document.Save.
    public static class DocumentExtensions
    {
        public static void Save(this Document doc, string fileName, CancellationToken cancellationToken)
        {
            // Use OoxmlSaveOptions for DOCX format and attach a progress callback that checks the token.
            var saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                ProgressCallback = new CancelSavingCallback(cancellationToken)
            };

            // Perform the save operation with the configured options.
            doc.Save(fileName, saveOptions);
        }
    }

    // Callback that aborts saving when the cancellation token is set.
    internal class CancelSavingCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _cancellationToken;

        public CancelSavingCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
        }

        public void Notify(DocumentSavingArgs args)
        {
            if (_cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException("Document saving was cancelled.");
        }
    }

    class Program
    {
        static void Main()
        {
            // Prepare a simple document.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Hello, Aspose.Words with cancellation support!");

            // Define output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");

            // Create a cancellation token that is already cancelled to demonstrate aborting.
            var cts = new CancellationTokenSource();
            cts.Cancel(); // Immediate cancellation.

            try
            {
                // Attempt to save using the extension method.
                doc.Save(outputPath, cts.Token);
                Console.WriteLine("Document saved successfully.");
            }
            catch (OperationCanceledException ex)
            {
                Console.WriteLine($"Save operation cancelled: {ex.Message}");
            }

            // Verify that the file does not exist when cancelled.
            if (File.Exists(outputPath))
            {
                Console.WriteLine("File was created despite cancellation.");
            }
            else
            {
                Console.WriteLine("No file was created, as expected.");
            }
        }
    }
}
