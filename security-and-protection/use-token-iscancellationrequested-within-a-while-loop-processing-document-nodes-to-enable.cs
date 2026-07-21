using System;
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
        for (int p = 1; p <= 10; p++)
        {
            builder.Writeln($"Paragraph {p}");
        }

        // Save the original document to satisfy the save requirement.
        const string sourcePath = "Source.docx";
        doc.Save(sourcePath);

        // Set up a cancellation token source.
        CancellationTokenSource cts = new CancellationTokenSource();

        // Simulate an external cancellation request after a short delay.
        Task.Run(async () =>
        {
            await Task.Delay(200); // 200 milliseconds
            cts.Cancel();
        });

        // Process all nodes in the document, checking for cancellation on each iteration.
        NodeCollection allNodes = doc.GetChildNodes(NodeType.Any, true);
        int index = 0;
        while (index < allNodes.Count)
        {
            // Gracefully exit the loop if cancellation is requested.
            if (cts.Token.IsCancellationRequested)
            {
                Console.WriteLine("Cancellation requested. Exiting processing loop.");
                break;
            }

            // Example processing: output the node type.
            Node node = allNodes[index];
            Console.WriteLine($"Processing node {index + 1}/{allNodes.Count}: {node.NodeType}");

            index++;
        }

        // Save the (potentially partially processed) document.
        const string outputPath = "Processed.docx";
        doc.Save(outputPath);
    }
}
