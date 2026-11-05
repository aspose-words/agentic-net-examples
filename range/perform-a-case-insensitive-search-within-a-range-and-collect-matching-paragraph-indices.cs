using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    // Callback that records paragraph indices containing a match.
    private class CollectParagraphsCallback : IReplacingCallback
    {
        private readonly NodeCollection _allParagraphs;
        public HashSet<int> MatchingIndices { get; } = new HashSet<int>();

        public CollectParagraphsCallback(NodeCollection allParagraphs)
        {
            _allParagraphs = allParagraphs;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
        {
            // The match node is a Run; get its containing paragraph.
            Paragraph paragraph = (Paragraph)e.MatchNode.GetAncestor(NodeType.Paragraph);
            if (paragraph != null)
            {
                int index = _allParagraphs.IndexOf(paragraph);
                if (index >= 0)
                    MatchingIndices.Add(index);
            }

            // Skip actual replacement; we only want to collect matches.
            return ReplaceAction.Skip;
        }
    }

    public static void Main()
    {
        // Create a new document and add sample paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph with Lorem ipsum.");
        builder.Writeln("Second paragraph without the keyword.");
        builder.Writeln("Third paragraph with lorem again.");
        builder.Writeln("Fourth paragraph with LOREM in uppercase.");

        // Save the sample document (demonstrates lifecycle).
        doc.Save("Sample.docx");

        // Define the search term (case‑insensitive).
        string searchTerm = "lorem";

        // Get all paragraphs in the document for index lookup.
        NodeCollection allParagraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        // Set up find/replace options for a case‑insensitive search.
        FindReplaceOptions findOptions = new FindReplaceOptions
        {
            MatchCase = false // case‑insensitive
        };

        // Use a callback to collect paragraph indices without modifying the document.
        var collector = new CollectParagraphsCallback(allParagraphs);
        findOptions.ReplacingCallback = collector;

        // Perform a replace operation that actually skips replacement; this triggers the callback.
        doc.Range.Replace(searchTerm, "", findOptions);

        // Output the collected paragraph indices.
        Console.WriteLine($"Paragraph indices containing the term \"{searchTerm}\":");
        foreach (int idx in collector.MatchingIndices)
        {
            Console.WriteLine(idx);
        }
    }
}
