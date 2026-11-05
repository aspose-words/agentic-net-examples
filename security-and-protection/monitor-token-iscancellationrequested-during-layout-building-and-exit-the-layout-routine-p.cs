using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world. This is a test document.");

        // Save the initial document (optional, just to have a file on disk).
        const string initialPath = "initial.docx";
        doc.Save(initialPath);

        // Set up a cancellation token source.
        CancellationTokenSource cts = new CancellationTokenSource();

        // Start layout building in a separate task.
        Task layoutTask = Task.Run(() =>
        {
            // Check for cancellation before starting the layout operation.
            if (cts.Token.IsCancellationRequested)
                throw new OperationCanceledException(cts.Token);

            // Aspose.Words does not accept a cancellation token directly,
            // so we perform the layout only if cancellation has not been requested.
            doc.UpdatePageLayout();

            // After the layout finishes, check again in case cancellation was requested
            // during the operation (unlikely for a single call, but follows the requirement).
            if (cts.Token.IsCancellationRequested)
                throw new OperationCanceledException(cts.Token);
        }, cts.Token);

        // Simulate a condition that triggers cancellation shortly after the task starts.
        Thread.Sleep(100);
        cts.Cancel();

        try
        {
            layoutTask.Wait();
            Console.WriteLine("Layout completed without cancellation.");
        }
        catch (AggregateException ex) when (ex.InnerException is OperationCanceledException)
        {
            Console.WriteLine("Layout was cancelled as requested.");
        }

        // Save the document only if layout was not cancelled.
        if (!cts.IsCancellationRequested)
        {
            const string finalPath = "final.docx";
            doc.Save(finalPath);
            Console.WriteLine($"Document saved to '{finalPath}'.");
        }
        else
        {
            Console.WriteLine("Document not saved because layout was cancelled.");
        }
    }
}
