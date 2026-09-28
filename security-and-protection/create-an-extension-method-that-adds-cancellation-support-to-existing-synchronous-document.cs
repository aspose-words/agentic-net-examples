using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;

namespace AsposeWordsCancellationDemo
{
    public static class DocumentExtensions
    {
        // Adds cancellation support to the synchronous Document.Save method.
        public static async Task SaveAsync(this Document doc, string fileName, CancellationToken cancellationToken = default)
        {
            // Throw if cancellation was requested before starting the operation.
            cancellationToken.ThrowIfCancellationRequested();

            // Execute the blocking Save on a thread‑pool thread.
            await Task.Run(() =>
            {
                // Check again before invoking the save.
                cancellationToken.ThrowIfCancellationRequested();
                doc.Save(fileName);
            }, cancellationToken);
        }
    }

    public class Program
    {
        public static async Task Main()
        {
            // Create a simple document in memory.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello, Aspose.Words with cancellation support!");

            // Define a temporary output file path.
            string outputPath = Path.Combine(Path.GetTempPath(), "CancellationDemo.docx");

            // Ensure a clean start.
            if (File.Exists(outputPath))
                File.Delete(outputPath);

            // Use a cancellation token that is not cancelled.
            using CancellationTokenSource cts = new CancellationTokenSource();

            // Save the document using the extension method.
            await doc.SaveAsync(outputPath, cts.Token);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved as expected.");

            // Load the saved document to confirm it can be opened.
            Document loaded = new Document(outputPath);
            // No further actions needed; successful load confirms the save.

            // Optional cleanup.
            // File.Delete(outputPath);
        }
    }
}
