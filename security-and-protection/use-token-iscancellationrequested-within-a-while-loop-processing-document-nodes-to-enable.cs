using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several paragraphs.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 10; i++)
        {
            builder.Writeln($"Paragraph {i}");
        }

        // Save the source document locally.
        string sourcePath = "Source.docx";
        doc.Save(sourcePath);

        // Load the document for processing.
        var loadDoc = new Document(sourcePath);

        // Set up a cancellation token that will be triggered after a short delay.
        var cts = new CancellationTokenSource();
        Task.Delay(200).ContinueWith(_ => cts.Cancel());

        // Process document nodes in a while loop, checking for cancellation.
        NodeCollection nodes = loadDoc.GetChildNodes(NodeType.Any, true);
        int index = 0;
        while (index < nodes.Count)
        {
            if (cts.Token.IsCancellationRequested)
            {
                // Graceful exit when cancellation is requested.
                Console.WriteLine("Cancellation requested. Exiting processing loop.");
                break;
            }

            // Example processing: prepend text to each paragraph.
            var node = nodes[index];
            if (node.NodeType == NodeType.Paragraph)
            {
                var paragraph = (Paragraph)node;
                paragraph.PrependChild(new Run(loadDoc, "[Processed] "));
            }

            index++;
        }

        // Save the processed document.
        string outputPath = "Processed.docx";
        loadDoc.Save(outputPath);

        // Validate that the output file exists.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Processed document was not saved.");
    }
}
