using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;

namespace AsposeWordsCancellationExample
{
    // Extension methods for Aspose.Words.Document
    public static class DocumentExtensions
    {
        /// <summary>
        /// Saves the document to the specified file path with support for cancellation.
        /// The method checks the cancellation token before starting the save operation.
        /// If the token is cancelled during the save, an OperationCanceledException is thrown.
        /// </summary>
        public static void SaveWithCancellation(this Document document, string filePath, CancellationToken cancellationToken)
        {
            // Throw if cancellation was already requested.
            cancellationToken.ThrowIfCancellationRequested();

            // Run the synchronous Save method on a background thread.
            // This allows the cancellation token to be observed while the operation is pending.
            Task saveTask = Task.Run(() => document.Save(filePath), cancellationToken);

            try
            {
                // Wait for the save to complete, propagating cancellation if it occurs.
                saveTask.Wait(cancellationToken);
            }
            catch (AggregateException ae)
            {
                // Unwrap the inner exception if it is a cancellation.
                if (ae.InnerException is OperationCanceledException)
                    throw ae.InnerException;
                throw;
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a simple document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello, Aspose.Words with cancellation support!");

            // Define the output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CancelledSaveExample.docx");

            // Create a cancellation token source (not cancelled in this example).
            using (CancellationTokenSource cts = new CancellationTokenSource())
            {
                // Save the document using the extension method.
                doc.SaveWithCancellation(outputPath, cts.Token);
            }

            // Validate that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved as expected.");

            // Optionally, load the saved document to ensure it is readable.
            Document loadedDoc = new Document(outputPath);
            Console.WriteLine("Document saved and loaded successfully. Text content:");
            Console.WriteLine(loadedDoc.GetText().Trim());
        }
    }
}
