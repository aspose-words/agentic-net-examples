using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with many lines to make layout take noticeable time.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample Document");
        for (int i = 0; i < 5000; i++)
        {
            builder.Writeln($"Line {i + 1}: The quick brown fox jumps over the lazy dog.");
        }

        // Set up a cancellation token that will be triggered after a short delay.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Cancel after 100 milliseconds.
            Task.Delay(100).ContinueWith(_ => cts.Cancel());

            try
            {
                // Check for cancellation before starting the layout operation.
                cts.Token.ThrowIfCancellationRequested();

                // Aspose.Words does not provide a direct overload with CancellationToken,
                // so we perform the layout synchronously and rely on the pre‑check.
                doc.UpdatePageLayout();

                // If layout completes, save the document.
                string outputPath = "output.docx";
                doc.Save(outputPath);

                // Verify the file was created.
                if (!File.Exists(outputPath))
                    throw new InvalidOperationException("The document was not saved as expected.");

                Console.WriteLine("Layout completed and document saved successfully.");
            }
            catch (OperationCanceledException)
            {
                // Layout was cancelled before it started.
                Console.WriteLine("Layout operation was cancelled via the cancellation token.");
            }
        }
    }
}
