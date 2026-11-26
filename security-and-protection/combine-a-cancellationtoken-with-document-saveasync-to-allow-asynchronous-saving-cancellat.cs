using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static async Task Main()
    {
        // Create a simple document with one paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words async save with cancellation!");

        // Define the output file path.
        string outputPath = "AsyncCancelled.docx";

        // Ensure any previous file is removed.
        if (File.Exists(outputPath))
            File.Delete(outputPath);

        // Set up a CancellationTokenSource that cancels after 100 milliseconds.
        using CancellationTokenSource cts = new CancellationTokenSource(100);

        try
        {
            // Perform the save operation on a background thread, passing the cancellation token.
            await Task.Run(() =>
            {
                // Throw if cancellation was requested before the save starts.
                cts.Token.ThrowIfCancellationRequested();

                // Synchronous save – Aspose.Words does not provide a native SaveAsync method.
                doc.Save(outputPath, new OoxmlSaveOptions());

                // Optionally, check cancellation again after the save (if the operation were longer).
                cts.Token.ThrowIfCancellationRequested();
            }, cts.Token);

            Console.WriteLine("Document saved successfully (unexpected).");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Save operation was cancelled as expected.");
        }

        // Check whether a file was created and clean up if necessary.
        if (File.Exists(outputPath))
        {
            File.Delete(outputPath);
            Console.WriteLine("Partial file removed.");
        }
        else
        {
            Console.WriteLine("No file was created.");
        }
    }
}
