using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with many paragraphs to make layout processing noticeable.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 5000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        // Prepare a cancellation token that will be triggered after a short delay.
        using CancellationTokenSource cts = new CancellationTokenSource();
        // Cancel after 100 milliseconds.
        Task.Run(async () =>
        {
            await Task.Delay(100);
            cts.Cancel();
        });

        // Build the layout while monitoring the cancellation token.
        // Since Aspose.Words does not provide an overload with CancellationToken,
        // run the layout operation on a separate task and wait with the token.
        Task layoutTask = Task.Run(() => doc.UpdatePageLayout(), cts.Token);
        try
        {
            layoutTask.Wait(cts.Token);
            Console.WriteLine("Layout completed without cancellation.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Layout was cancelled via the cancellation token.");
            return;
        }
        catch (AggregateException ae) when (ae.InnerExceptions.Any(e => e is OperationCanceledException))
        {
            Console.WriteLine("Layout was cancelled via the cancellation token.");
            return;
        }

        // Save the document if layout completed.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "LayoutResult.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved as expected.");

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
