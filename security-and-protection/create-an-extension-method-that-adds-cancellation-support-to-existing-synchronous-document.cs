using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;

namespace AsposeWordsCancellationDemo
{
    // Extension methods for Aspose.Words Document to add cancellation support.
    public static class DocumentExtensions
    {
        // Saves the document to a file with cancellation support.
        public static Task SaveAsync(this Document document, string fileName, CancellationToken cancellationToken = default)
        {
            // Run the synchronous Save inside a Task so the token can be observed.
            return Task.Run(() =>
            {
                // Throw if cancellation was requested before starting the save.
                cancellationToken.ThrowIfCancellationRequested();

                // Perform the actual save operation.
                document.Save(fileName);
            }, cancellationToken);
        }
    }

    public class Program
    {
        public static async Task Main(string[] args)
        {
            // Create a simple document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello, Aspose.Words with cancellation support!");

            string outputPath = "output.docx";

            // Normal save with cancellation token that is not cancelled.
            using (CancellationTokenSource cts = new CancellationTokenSource())
            {
                await doc.SaveAsync(outputPath, cts.Token);
            }

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"The file '{outputPath}' was not created.");

            // Demonstrate cancellation before the save starts.
            using (CancellationTokenSource cts = new CancellationTokenSource())
            {
                cts.Cancel(); // Cancel immediately.

                try
                {
                    await doc.SaveAsync("should_not_be_created.docx", cts.Token);
                    throw new InvalidOperationException("SaveAsync should have been cancelled but completed.");
                }
                catch (OperationCanceledException)
                {
                    // Expected cancellation.
                }
            }

            // Clean up created files.
            File.Delete(outputPath);
        }
    }
}
