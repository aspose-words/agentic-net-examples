using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World");
        builder.Writeln("This is a sample paragraph.");
        builder.Writeln("Another line with hello.");
        builder.Writeln("No match here.");
        builder.Writeln("HELLO again!");

        // Term to search for (case‑insensitive).
        string searchTerm = "hello";

        // Collector for paragraph indices that contain the term.
        var collector = new ParagraphIndexCollector(doc);

        // Configure case‑insensitive search and attach the callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            ReplacingCallback = collector   // Use the callback via options.
        };

        // Perform the search. No actual replacement occurs because the callback returns Skip.
        doc.Range.Replace(searchTerm, string.Empty, options);

        // Output the collected paragraph indices.
        Console.WriteLine($"Paragraph indices containing the term \"{searchTerm}\":");
        foreach (int index in collector.ParagraphIndices)
        {
            Console.WriteLine(index);
        }
    }

    // Callback that records the index of each paragraph where a match is found.
    private class ParagraphIndexCollector : IReplacingCallback
    {
        private readonly Document _document;
        public List<int> ParagraphIndices { get; } = new List<int>();

        public ParagraphIndexCollector(Document document)
        {
            _document = document;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
        {
            // The match node is a Run; get its parent paragraph.
            if (e.MatchNode is Run run)
            {
                Paragraph paragraph = run.ParentParagraph;
                NodeCollection allParagraphs = _document.GetChildNodes(NodeType.Paragraph, true);
                int index = allParagraphs.IndexOf(paragraph);
                if (index >= 0 && !ParagraphIndices.Contains(index))
                {
                    ParagraphIndices.Add(index);
                }
            }

            // Skip actual replacement.
            return ReplaceAction.Skip;
        }
    }
}
