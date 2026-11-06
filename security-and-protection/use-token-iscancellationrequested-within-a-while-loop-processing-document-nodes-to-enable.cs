using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 10; i++)
        {
            builder.Writeln($"Paragraph {i}");
        }

        // Save the source document locally.
        string sourcePath = "source.docx";
        doc.Save(sourcePath);

        // Prepare a cancellation token that will be triggered after a short delay.
        using CancellationTokenSource cts = new CancellationTokenSource();
        CancellationToken token = cts.Token;

        // Cancel after 200 milliseconds to simulate a graceful exit request.
        Task.Run(async () =>
        {
            await Task.Delay(200);
            cts.Cancel();
        });

        // Process document nodes in a while loop, checking for cancellation.
        NodeCollection nodes = doc.GetChildNodes(NodeType.Any, true);
        int index = 0;
        while (index < nodes.Count)
        {
            if (token.IsCancellationRequested)
            {
                Console.WriteLine("Cancellation requested. Exiting processing loop.");
                break;
            }

            // Example processing: output node type.
            Node node = nodes[index];
            Console.WriteLine($"Processing node {index + 1}/{nodes.Count}: {node.NodeType}");

            // Simulate some work.
            Thread.Sleep(50);

            index++;
        }

        // Save the (potentially partially processed) document.
        string outputPath = "processed.docx";
        doc.Save(outputPath);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"The output file '{outputPath}' was not created.");
        }

        Console.WriteLine("Processing completed. Output saved to " + outputPath);
    }
}
