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

        builder.Writeln("This is the first paragraph containing Aspose.Words.");
        builder.Writeln("Second paragraph without the keyword.");
        builder.Writeln("Third paragraph mentions aspose in lower case.");
        builder.Writeln("Another line with the word ASPose in mixed case.");
        builder.Writeln("Final paragraph without it.");

        // The term to search for (case‑insensitive).
        string searchTerm = "aspose";

        // List to collect zero‑based paragraph indices where matches are found.
        List<int> matchingParagraphIndices = new List<int>();

        // Set up find/replace options for a case‑insensitive search.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false // ignore case
        };
        options.ReplacingCallback = new SearchCallback(matchingParagraphIndices, doc);

        // Perform a replace where the replacement text is identical to the pattern.
        // The callback records matches and skips actual replacement.
        doc.Range.Replace(searchTerm, searchTerm, options);

        // Output the collected paragraph indices.
        Console.WriteLine("Paragraph indices containing the term \"{0}\":", searchTerm);
        foreach (int index in matchingParagraphIndices)
        {
            Console.WriteLine(index);
        }
    }

    // Callback that records the paragraph index of each match and skips replacement.
    private class SearchCallback : IReplacingCallback
    {
        private readonly List<int> _indices;
        private readonly Document _document;

        public SearchCallback(List<int> indices, Document document)
        {
            _indices = indices;
            _document = document;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Find the paragraph that contains the matched run.
            Node matchNode = args.MatchNode;
            Paragraph paragraph = (Paragraph)matchNode.GetAncestor(NodeType.Paragraph);

            // Determine the paragraph's index within the body.
            int paragraphIndex = _document.FirstSection.Body.Paragraphs.IndexOf(paragraph);

            // Record the index if it hasn't been recorded yet (multiple matches in same paragraph).
            if (!_indices.Contains(paragraphIndex))
                _indices.Add(paragraphIndex);

            // Skip actual replacement to keep the document unchanged.
            return ReplaceAction.Skip;
        }
    }
}
