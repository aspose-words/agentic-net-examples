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
        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document will be saved asynchronously with cancellation support.");

        string outputPath = Path.Combine(Path.GetTempPath(), "AsyncCancelled.docx");

        // Remove any existing file.
        if (File.Exists(outputPath))
            File.Delete(outputPath);

        // Set up a cancellation token that will be triggered shortly.
        using CancellationTokenSource cts = new CancellationTokenSource();

        // Cancel after 100 milliseconds.
        Task cancelTask = Task.Delay(100).ContinueWith(_ => cts.Cancel());

        try
        {
            // Perform the save operation on a background thread, passing the token.
            await Task.Run(() =>
            {
                // Throw if cancellation was already requested.
                cts.Token.ThrowIfCancellationRequested();

                // Save the document synchronously (Aspose.Words does not provide SaveAsync).
                doc.Save(outputPath, new OoxmlSaveOptions(SaveFormat.Docx));
            }, cts.Token);

            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Save operation was cancelled.");
            // Clean up partially written file if it exists.
            if (File.Exists(outputPath))
                File.Delete(outputPath);
        }

        // Ensure the cancellation task completes.
        await cancelTask;
    }
}
