using System;
using System.Threading;
using Aspose.Words;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Simple argument validation (input, output paths optional).
            string inputPath = args.Length > 0 ? args[0] : "input.docx";
            string outputPath = args.Length > 1 ? args[1] : "output.docx";

            // Create a cancellation token source – in a real app you could hook this to a UI button, Ctrl‑C, etc.
            using var cts = new CancellationTokenSource();
            // For demonstration we do not request cancellation, but the token is ready for use.

            var processor = new DocumentProcessor();
            processor.ProcessDocument(inputPath, outputPath, cts.Token);

            Console.WriteLine($"Document processed. Saved to '{outputPath}'.");
        }
    }

    public class DocumentProcessor
    {
        /// <summary>
        /// Loads a Word document, processes its nodes in a loop that can be cancelled,
        /// and saves the result.
        /// </summary>
        /// <param name="inputPath">Path to the source .docx file.</param>
        /// <param name="outputPath">Path where the processed file will be saved.</param>
        /// <param name="token">Cancellation token to allow graceful exit.</param>
        public void ProcessDocument(string inputPath, string outputPath, CancellationToken token)
        {
            // Load the document (lifecycle rule: load)
            Document doc = new Document(inputPath);

            // Begin traversing the nodes of the main body.
            Node curNode = doc.FirstSection.Body.FirstChild;

            while (curNode != null)
            {
                // Check for cancellation request on each iteration.
                if (token.IsCancellationRequested)
                {
                    // Exit the loop without throwing; the document will be saved in its current state.
                    break;
                }

                // Store reference to the next sibling before any modification.
                Node nextNode = curNode.NextSibling;

                // Example processing: remove empty paragraphs.
                if (curNode.NodeType == NodeType.Paragraph && curNode.GetText().Trim() == string.Empty)
                {
                    curNode.Remove();
                }

                // Move to the next node.
                curNode = nextNode;
            }

            // Save the (potentially modified) document (lifecycle rule: save)
            doc.Save(outputPath);
        }
    }
}
