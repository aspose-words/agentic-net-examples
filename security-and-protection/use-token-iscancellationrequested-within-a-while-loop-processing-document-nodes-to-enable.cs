using System;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCancellationDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add several paragraphs to have nodes to process.
            for (int i = 1; i <= 20; i++)
            {
                builder.Writeln($"Paragraph {i}");
            }

            // Set up a cancellation token that will be triggered after a short delay.
            using (CancellationTokenSource cts = new CancellationTokenSource())
            {
                // Cancel after 200 milliseconds.
                cts.CancelAfter(200);
                CancellationToken token = cts.Token;

                // Get all nodes in the document (including nested nodes).
                var allNodes = doc.GetChildNodes(NodeType.Any, true);
                int index = 0;

                // Process nodes in a while loop, checking for cancellation.
                while (index < allNodes.Count)
                {
                    // Gracefully exit if cancellation is requested.
                    if (token.IsCancellationRequested)
                    {
                        Console.WriteLine("Cancellation requested. Exiting processing loop.");
                        break;
                    }

                    var node = allNodes[index];

                    // Example processing: append a marker to each paragraph.
                    if (node.NodeType == NodeType.Paragraph)
                    {
                        Paragraph paragraph = (Paragraph)node;
                        paragraph.AppendChild(new Run(doc, " [processed]"));
                    }

                    index++;
                }
            }

            // Save the resulting document.
            const string outputPath = "Processed.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to '{outputPath}'.");
        }
    }
}
