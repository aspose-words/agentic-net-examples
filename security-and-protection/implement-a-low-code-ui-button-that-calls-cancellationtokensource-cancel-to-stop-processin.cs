using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document and save it locally.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");
        string docPath = "sample.docx";
        doc.Save(docPath); // Save using the standard Save method.

        // Set up a cancellation token source that will be used to stop processing.
        using CancellationTokenSource cts = new CancellationTokenSource();

        // Start a background task that repeatedly loads and processes the document.
        Task processingTask = Task.Run(() => ProcessDocumentLoop(docPath, cts.Token));

        // Simulate a low‑code UI button click after a short delay.
        // In a real UI this would be an event handler; here we just delay and then cancel.
        Task.Delay(500).ContinueWith(_ => cts.Cancel());

        try
        {
            // Wait for the processing task to complete (it will end when cancelled).
            processingTask.Wait();
        }
        catch (AggregateException ae)
        {
            // Expect a TaskCanceledException wrapped in AggregateException.
            foreach (var ex in ae.InnerExceptions)
            {
                if (ex is OperationCanceledException) continue;
                Console.Error.WriteLine($"Unexpected exception: {ex}");
            }
        }

        // Clean up the sample file.
        if (File.Exists(docPath))
            File.Delete(docPath);
    }

    // Continuously loads the document and performs a trivial operation until cancelled.
    private static void ProcessDocumentLoop(string path, CancellationToken token)
    {
        while (true)
        {
            token.ThrowIfCancellationRequested();

            // Load the document using the constructor that takes a file name.
            Document loadedDoc = new Document(path);

            // Perform a simple operation: count words.
            string text = loadedDoc.GetText();
            int wordCount = text.Split(new[] { ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Length;

            // Simulate work by sleeping briefly.
            Thread.Sleep(100);
        }
    }
}
