using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace RangeSearchExample
{
    // Callback that collects paragraph indices of each match.
    class CollectingCallback : IReplacingCallback
    {
        private readonly Document _document;
        public List<int> ParagraphIndices { get; } = new List<int>();

        public CollectingCallback(Document document)
        {
            _document = document;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
        {
            // The node that contains the match (usually a Run).
            Node matchNode = e.MatchNode;

            // Find the ancestor paragraph of the match.
            Paragraph paragraph = (Paragraph)matchNode.GetAncestor(NodeType.Paragraph);
            if (paragraph != null)
            {
                // Get all paragraphs in the document.
                NodeCollection allParagraphs = _document.GetChildNodes(NodeType.Paragraph, true);
                // Determine the index of the paragraph.
                int index = allParagraphs.IndexOf(paragraph);
                ParagraphIndices.Add(index);
            }

            // Skip actual replacement.
            return ReplaceAction.Skip;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a sample document with several paragraphs.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("This is the first paragraph.");
            builder.Writeln("Aspose.Words is a powerful library.");
            builder.Writeln("The word aspose appears here in lower case.");
            builder.Writeln("Another line without the keyword.");
            builder.Writeln("ASPOSE can be written in upper case as well.");
            builder.Writeln("Final paragraph.");

            // Save the sample document (optional, demonstrates lifecycle rule).
            string filePath = "Sample.docx";
            doc.Save(filePath);

            // Define the search term (case‑insensitive).
            string searchTerm = "aspose";

            // Set up find/replace options for a case‑insensitive search.
            FindReplaceOptions options = new FindReplaceOptions
            {
                MatchCase = false
            };

            // Create the callback that will collect paragraph indices.
            CollectingCallback callback = new CollectingCallback(doc);
            options.ReplacingCallback = callback;

            // Perform a replace operation that does not modify the document
            // but invokes the callback for each match.
            doc.Range.Replace(searchTerm, "", options);

            // Output the collected paragraph indices.
            Console.WriteLine($"Paragraph indices containing the term \"{searchTerm}\":");
            foreach (int index in callback.ParagraphIndices)
            {
                Console.WriteLine(index);
            }
        }
    }
}
