using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCancellationDemo
{
    public class Program
    {
        public static void Main()
        {
            // Prepare a cancellation token that will be triggered after a short delay.
            using var cts = new CancellationTokenSource();
            cts.CancelAfter(200); // Cancel after 200 milliseconds.

            // Create a sample document and add several paragraphs.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            for (int p = 1; p <= 20; p++)
            {
                builder.Writeln($"Paragraph {p}");
            }

            // Save the original document.
            string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docx");
            doc.Save(originalPath);

            // Process document nodes while respecting the cancellation token.
            var allNodes = doc.GetChildNodes(NodeType.Any, true);
            int index = 0;
            while (index < allNodes.Count && !cts.Token.IsCancellationRequested)
            {
                var node = allNodes[index];

                // Example modification: if the node is a paragraph, append a marker.
                if (node.NodeType == NodeType.Paragraph)
                {
                    var paragraph = (Paragraph)node;
                    paragraph.AppendChild(new Run(doc, " [Processed]"));
                }

                index++;
            }

            // Save the processed document.
            string processedPath = Path.Combine(Directory.GetCurrentDirectory(), "Processed.docx");
            doc.Save(processedPath);

            // Validate that the output file was created.
            if (!File.Exists(processedPath))
                throw new InvalidOperationException("The processed document was not saved correctly.");

            // Indicate completion (no interactive input required).
            Console.WriteLine("Document processing completed.");
        }
    }
}
