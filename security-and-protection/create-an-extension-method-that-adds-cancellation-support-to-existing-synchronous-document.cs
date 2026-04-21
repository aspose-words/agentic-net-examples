using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsCancellationDemo
{
    // Extension methods for Document.
    public static class DocumentExtensions
    {
        // Saves a document with cancellation support.
        public static void SaveWithCancellation(this Document doc, string fileName, CancellationToken cancellationToken)
        {
            // Create save options appropriate for the file extension.
            SaveOptions saveOptions = SaveOptions.CreateSaveOptions(Path.GetExtension(fileName));
            // Attach a progress callback that checks the cancellation token.
            saveOptions.ProgressCallback = new CancellationSavingCallback(cancellationToken);
            // Perform the save operation.
            doc.Save(fileName, saveOptions);
        }

        // Callback that aborts saving when cancellation is requested.
        private class CancellationSavingCallback : IDocumentSavingCallback
        {
            private readonly CancellationToken _cancellationToken;

            public CancellationSavingCallback(CancellationToken cancellationToken)
            {
                _cancellationToken = cancellationToken;
            }

            public void Notify(DocumentSavingArgs args)
            {
                if (_cancellationToken.IsCancellationRequested)
                {
                    throw new OperationCanceledException(
                        $"Saving aborted at {args.EstimatedProgress}% progress.");
                }
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // Example 1: Normal save (no cancellation).
            // -----------------------------------------------------------------
            string normalPath = "NormalSave.docx";
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello, Aspose.Words with cancellation support!");

            // No cancellation requested.
            using (CancellationTokenSource cts = new CancellationTokenSource())
            {
                doc.SaveWithCancellation(normalPath, cts.Token);
            }

            // Verify the file was created.
            if (!File.Exists(normalPath))
                throw new Exception($"Failed to create '{normalPath}'.");

            // -----------------------------------------------------------------
            // Example 2: Save that is cancelled immediately.
            // -----------------------------------------------------------------
            string cancelledPath = "CancelledSave.docx";
            Document largeDoc = new Document();
            DocumentBuilder largeBuilder = new DocumentBuilder(largeDoc);
            for (int i = 0; i < 1000; i++)
                largeBuilder.Writeln($"Line {i + 1}");

            // Cancel before the save starts.
            using (CancellationTokenSource cts = new CancellationTokenSource())
            {
                cts.Cancel(); // Request cancellation.

                try
                {
                    largeDoc.SaveWithCancellation(cancelledPath, cts.Token);
                }
                catch (OperationCanceledException ex)
                {
                    Console.WriteLine($"Save operation was cancelled: {ex.Message}");
                }
            }

            // Ensure the cancelled file does not exist.
            if (File.Exists(cancelledPath))
                File.Delete(cancelledPath);
        }
    }
}
