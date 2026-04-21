using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Paragraph {i}");
        }

        // Save the sample document locally.
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        doc.Save(sourcePath);

        // Load the document for processing.
        Document loadedDoc = new Document(sourcePath);

        // Set up a cancellation token source.
        CancellationTokenSource cts = new CancellationTokenSource();
        CancellationToken token = cts.Token;

        // Retrieve all paragraph nodes in the document.
        NodeCollection paragraphs = loadedDoc.GetChildNodes(NodeType.Paragraph, true);
        int index = 0;

        // Process paragraphs in a while loop, checking for cancellation.
        while (index < paragraphs.Count)
        {
            // Gracefully exit if cancellation has been requested.
            if (token.IsCancellationRequested)
            {
                Console.WriteLine("Cancellation requested. Exiting processing loop.");
                break;
            }

            // Example processing: append a marker to each paragraph.
            Paragraph para = (Paragraph)paragraphs[index];
            para.AppendChild(new Run(loadedDoc, " - processed"));

            index++;

            // Simulate a condition that triggers cancellation after two paragraphs.
            if (index == 2)
            {
                cts.Cancel();
            }
        }

        // Save the modified document.
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        loadedDoc.Save(resultPath);

        // Validate that the output file was created.
        if (!File.Exists(resultPath))
        {
            throw new InvalidOperationException("The result document was not saved correctly.");
        }

        // Optional: indicate completion (no user interaction required).
        Console.WriteLine("Processing completed.");
    }
}
