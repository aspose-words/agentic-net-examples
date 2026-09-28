using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple document with several paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int p = 1; p <= 5; p++)
        {
            builder.Writeln($"Paragraph {p}");
        }

        // Prepare a cancellation token source.
        CancellationTokenSource cts = new CancellationTokenSource();

        // Simulate external cancellation after a short delay.
        Task.Run(() =>
        {
            Thread.Sleep(100); // Cancel after 100 ms.
            cts.Cancel();
        });

        // Process document nodes in a while loop, checking for cancellation.
        NodeCollection nodes = doc.GetChildNodes(NodeType.Any, true);
        int index = 0;
        while (index < nodes.Count)
        {
            if (cts.Token.IsCancellationRequested)
            {
                Console.WriteLine("Cancellation requested. Exiting processing loop.");
                break;
            }

            // Example processing: output node type.
            Node node = nodes[index];
            Console.WriteLine($"Processing node {index + 1}/{nodes.Count}: {node.NodeType}");
            index++;
        }

        // Save the (potentially partially processed) document.
        string outputPath = "Processed.docx";
        doc.Save(outputPath);

        // Validate that the file was saved.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create output file: {outputPath}");
        }

        Console.WriteLine($"Document saved to {Path.GetFullPath(outputPath)}");
    }
}
