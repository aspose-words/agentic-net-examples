using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Callback that checks the cancellation token and aborts saving when requested.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _token;

        public SavingProgressCallback(CancellationToken token) => _token = token;

        public void Notify(DocumentSavingArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException($"Saving canceled at {args.EstimatedProgress}% progress.");
        }
    }

    public static async Task Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "LargeDocument.docx");

        // Create a large document to make the save operation take noticeable time.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 5000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        // Set up a cancellation token that will be triggered after a short delay.
        using var cts = new CancellationTokenSource();
        cts.CancelAfter(200); // Cancel after 200 milliseconds.

        try
        {
            // Configure save options with a progress callback that respects the token.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                ProgressCallback = new SavingProgressCallback(cts.Token)
            };

            // Perform the save operation. The callback will throw if cancellation is requested.
            await Task.Run(() => doc.Save(outputPath, saveOptions), cts.Token);
            Console.WriteLine("Document saved successfully.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Document saving was canceled.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }

        // Verify whether the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Output file exists at: {outputPath}");
        }
        else
        {
            Console.WriteLine("Output file does not exist (save was likely canceled).");
        }
    }
}
