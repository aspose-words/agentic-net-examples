using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with many paragraphs to make the save operation take noticeable time.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 2000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        // Set up a cancellation token that will be triggered after a short delay.
        using var cts = new CancellationTokenSource();
        cts.CancelAfter(200); // Cancel after 200 milliseconds.

        // Configure save options with a progress callback that monitors the token.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new SavingProgressCallback(cts.Token)
        };

        string outputPath = "output.docx";

        try
        {
            // Attempt to save the document. The callback may abort the operation.
            doc.Save(outputPath, saveOptions);
            Console.WriteLine("Document saved successfully.");
        }
        catch (OperationCanceledException ex)
        {
            // Inform the user that the operation was canceled and provide progress info.
            Console.WriteLine($"Saving was canceled: {ex.Message}");
        }

        // Verify whether the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"File exists: {Path.GetFullPath(outputPath)}");
        }
        else
        {
            Console.WriteLine("File was not created due to cancellation.");
        }
    }

    // Progress callback that checks the cancellation token and aborts if requested.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _token;

        public SavingProgressCallback(CancellationToken token)
        {
            _token = token;
        }

        public void Notify(DocumentSavingArgs args)
        {
            if (_token.IsCancellationRequested)
            {
                // Throwing this exception aborts the save operation.
                throw new OperationCanceledException(
                    $"EstimatedProgress = {args.EstimatedProgress}%; Save operation canceled.");
            }
        }
    }
}
