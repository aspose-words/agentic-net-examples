using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;

namespace AsposeWordsCancellationDemo
{
    // Extension method that adds cancellation support to Document.Save.
    public static class DocumentExtensions
    {
        public static async Task SaveAsync(this Document doc, string fileName, CancellationToken cancellationToken = default)
        {
            // Throw if cancellation requested before starting the save.
            cancellationToken.ThrowIfCancellationRequested();

            // Run the synchronous Save on a background thread, observing the token.
            await Task.Run(() =>
            {
                // Check again inside the task in case token was cancelled while queued.
                cancellationToken.ThrowIfCancellationRequested();
                doc.Save(fileName);
            }, cancellationToken);
        }
    }

    class Program
    {
        static async Task Main()
        {
            // Create a simple document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello, Aspose.Words with cancellation support!");

            // Define output path.
            string outputPath = Path.Combine(Path.GetTempPath(), "CancellationDemo.docx");

            // Ensure any existing file is removed.
            if (File.Exists(outputPath))
                File.Delete(outputPath);

            // Create a cancellation token source (not cancelled).
            using CancellationTokenSource cts = new CancellationTokenSource();

            // Save the document using the extension method.
            await doc.SaveAsync(outputPath, cts.Token);

            // Verify the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("Document was not saved.");

            // Load the saved document to confirm it is readable.
            Document loaded = new Document(outputPath);
            Console.WriteLine("Document saved successfully to: " + outputPath);
        }
    }
}
