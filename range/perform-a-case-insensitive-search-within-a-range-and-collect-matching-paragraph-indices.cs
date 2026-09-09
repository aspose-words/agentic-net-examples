using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    // Callback that records the index of each paragraph containing a match.
    private class MatchRecorder : IReplacingCallback
    {
        private readonly List<int> _paragraphIndices;

        public MatchRecorder(List<int> paragraphIndices)
        {
            _paragraphIndices = paragraphIndices;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Find the paragraph that contains the match.
            Node node = args.MatchNode;
            while (node != null && node.NodeType != NodeType.Paragraph)
                node = node.ParentNode;

            if (node is Paragraph paragraph)
            {
                // The paragraph's parent story is a Body (or HeaderFooter, etc.).
                // Retrieve the collection of paragraphs from that story.
                var body = paragraph.ParentNode as CompositeNode;
                if (body != null)
                {
                    var paragraphs = body.GetChildNodes(NodeType.Paragraph, true);
                    int index = paragraphs.IndexOf(paragraph);
                    _paragraphIndices.Add(index);
                }
            }

            // Skip actual replacement; we only want to record matches.
            return ReplaceAction.Skip;
        }
    }

    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Aspose.Words is a powerful library.");
        builder.Writeln("This example demonstrates a case‑insensitive search.");
        builder.Writeln("Search for the word 'example' in this document.");
        builder.Writeln("Another line without the keyword.");
        builder.Writeln("EXAMPLE appears in different case.");

        // Prepare a list to hold paragraph indices where the term is found.
        List<int> matchingParagraphIndices = new List<int>();

        // Configure find/replace options for a case‑insensitive search.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false, // case‑insensitive
            ReplacingCallback = new MatchRecorder(matchingParagraphIndices)
        };

        // Perform a find operation without modifying the document.
        // The pattern is the word we are searching for.
        string searchTerm = "example";
        doc.Range.Replace(searchTerm, searchTerm, options);

        // Output the collected paragraph indices.
        Console.WriteLine("Paragraph indices containing the term \"{0}\":", searchTerm);
        foreach (int idx in matchingParagraphIndices)
        {
            Console.WriteLine(idx);
        }

        // Save the document (optional, demonstrates the lifecycle rule).
        doc.Save("SampleDocument.docx");
    }
}
