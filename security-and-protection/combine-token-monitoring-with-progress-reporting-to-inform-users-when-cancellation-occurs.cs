using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        const int totalParagraphs = 5000;
        const string outputPath = "output.docx";

        // Create a cancellation token that will be triggered after a short delay.
        using CancellationTokenSource cts = new CancellationTokenSource();
        // Cancel after 200 ms.
        Timer cancelTimer = new Timer(_ => cts.Cancel(), null, 200, Timeout.Infinite);

        // Build the document while reporting progress and checking for cancellation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < totalParagraphs; i++)
        {
            // Write a paragraph.
            builder.Writeln($"Paragraph {i + 1}");

            // Report progress.
            int percentDone = (i + 1) * 100 / totalParagraphs;
            Console.WriteLine($"Building progress: {percentDone}%");

            // If cancellation was requested, stop building.
            if (cts.IsCancellationRequested)
            {
                Console.WriteLine("Cancellation requested during document building. Stopping.");
                break;
            }
        }

        // Dispose the timer – no longer needed.
        cancelTimer.Dispose();

        // If the document was fully built, attempt to save it.
        if (!cts.IsCancellationRequested)
        {
            try
            {
                // Save the document; no special progress callback is needed here.
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
                doc.Save(outputPath, saveOptions);
                Console.WriteLine("Document saved successfully.");
            }
            catch (OperationCanceledException)
            {
                // This block is kept for completeness; Aspose.Words may throw this if it supports cancellation.
                Console.WriteLine("Save operation was canceled.");
            }
        }
        else
        {
            Console.WriteLine("Document was not saved due to cancellation.");
        }

        // Verify whether the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"File '{outputPath}' exists.");
        }
        else
        {
            Console.WriteLine($"File '{outputPath}' was not created.");
        }
    }
}
