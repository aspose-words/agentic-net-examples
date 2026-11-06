using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static async Task Main()
    {
        // Create a simple source document.
        string sourcePath = "sample.docx";
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Hello, Aspose.Words!");
        sourceDoc.Save(sourcePath);

        // Prepare load options (default).
        LoadOptions loadOptions = new LoadOptions();

        // Open a stream for loading the document.
        using (FileStream stream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
        {
            // Cancel the token before starting the load operation.
            CancellationTokenSource cts = new CancellationTokenSource();
            cts.Cancel();

            try
            {
                // Attempt to load the document with a cancelled token.
                Document loadedDoc = await Task.Run(() => new Document(stream, loadOptions), cts.Token);

                // If loading somehow succeeds, save the document to verify.
                string outputPath = "loaded.docx";
                loadedDoc.Save(outputPath);
                Console.WriteLine($"Document loaded and saved to {outputPath}");
            }
            catch (OperationCanceledException)
            {
                // Handle cancellation and perform cleanup.
                Console.WriteLine("Document load was cancelled. Performing cleanup.");
                string partialPath = "loaded.docx";
                if (File.Exists(partialPath))
                {
                    File.Delete(partialPath);
                }
            }
            catch (AggregateException aggEx) when (aggEx.InnerException is OperationCanceledException)
            {
                // Handle cancellation wrapped in an AggregateException.
                Console.WriteLine("Document load was cancelled (AggregateException). Performing cleanup.");
                string partialPath = "loaded.docx";
                if (File.Exists(partialPath))
                {
                    File.Delete(partialPath);
                }
            }
        }

        // Final cleanup: delete the source document.
        if (File.Exists(sourcePath))
        {
            File.Delete(sourcePath);
        }
    }
}
