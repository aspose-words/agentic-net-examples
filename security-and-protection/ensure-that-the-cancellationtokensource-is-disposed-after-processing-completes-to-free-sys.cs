using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path for the output document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ProtectedDocument.docx");

        // Create a new document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world! This document will be processed with a CancellationTokenSource.");

        // Save the document to disk.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create output file: {outputPath}");

        // Use a CancellationTokenSource to control a simulated processing task.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Start a task that pretends to process the document.
            Task processingTask = Task.Run(() => ProcessDocument(outputPath, cts.Token), cts.Token);

            // Wait for the task to finish, but limit the wait time.
            bool completed = processingTask.Wait(TimeSpan.FromSeconds(5));

            // If the task didn't finish in time, request cancellation.
            if (!completed)
                cts.Cancel();

            // Propagate any exceptions from the task.
            if (processingTask.IsFaulted)
                throw processingTask.Exception!;
        }

        // The using block ensures the CancellationTokenSource is disposed here.
    }

    private static void ProcessDocument(string path, CancellationToken token)
    {
        // Simulate work that periodically checks for cancellation.
        for (int i = 0; i < 10; i++)
        {
            token.ThrowIfCancellationRequested();
            Thread.Sleep(200); // Simulated work.
        }

        // Reload the document to demonstrate load workflow.
        Document loadedDoc = new Document(path);

        // Read the document text (no modification needed for this example).
        string text = loadedDoc.GetText();

        // Placeholder for any security-related validation if required.
    }
}
