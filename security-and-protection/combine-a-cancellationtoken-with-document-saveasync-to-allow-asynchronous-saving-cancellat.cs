using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Helper method that runs the synchronous Save inside a Task and observes the cancellation token.
    private static async Task SaveDocumentAsync(Document doc, string path, SaveOptions options, CancellationToken token)
    {
        // Run the save operation on a background thread so it can be cancelled before it starts.
        await Task.Run(() =>
        {
            // Throw if cancellation was already requested.
            token.ThrowIfCancellationRequested();

            // Perform the actual save.
            doc.Save(path, options);
        }, token);
    }

    public static async Task Main(string[] args)
    {
        // Create a sample document with many paragraphs to make the save operation take noticeable time.
        Document doc = new Document();
        for (int i = 0; i < 5000; i++)
        {
            Paragraph para = new Paragraph(doc);
            para.AppendChild(new Run(doc, $"Paragraph {i + 1}"));
            doc.FirstSection.Body.AppendChild(para);
        }

        // Define the output file path.
        string outputPath = Path.Combine(Path.GetTempPath(), "AsyncSave.docx");

        // Ensure any previous file is removed.
        if (File.Exists(outputPath))
            File.Delete(outputPath);

        // ---------- Normal asynchronous save (no cancellation) ----------
        await SaveDocumentAsync(doc, outputPath, new OoxmlSaveOptions(), CancellationToken.None);
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Document was not saved as expected.");

        // ---------- Asynchronous save with cancellation ----------
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Cancel the token after 10 milliseconds.
            cts.CancelAfter(10);

            try
            {
                // Attempt to save the document asynchronously with the cancellable token.
                await SaveDocumentAsync(doc, outputPath, new OoxmlSaveOptions(), cts.Token);
                // If the operation completes without cancellation, indicate success.
                Console.WriteLine("Document saved successfully (cancellation did not occur).");
            }
            catch (OperationCanceledException)
            {
                // Expected path when the token is cancelled during the save operation.
                Console.WriteLine("Document save was cancelled via CancellationToken.");
            }
        }

        // Clean up the temporary file.
        if (File.Exists(outputPath))
            File.Delete(outputPath);
    }
}
