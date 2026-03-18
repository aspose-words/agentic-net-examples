using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the document.
        Document doc = new Document("Input.docx");

        // Text to search for (case‑insensitive).
        string searchText = "example";

        // List that will hold the zero‑based paragraph indices where matches are found.
        List<int> matchingParagraphIndices = new List<int>();

        // Configure find/replace options.
        FindReplaceOptions options = new FindReplaceOptions
        {
            // false => case‑insensitive search.
            MatchCase = false,
            // Use a callback to capture each match.
            ReplacingCallback = new CollectParagraphIndicesCallback(doc, matchingParagraphIndices)
        };

        // Perform a find‑and‑replace where the replacement is the same as the found text.
        // This triggers the callback without altering the document content.
        doc.Range.Replace(searchText, searchText, options);

        // Output the collected indices (for demonstration purposes).
        Console.WriteLine("Paragraph indices containing the text \"{0}\":", searchText);
        foreach (int idx in matchingParagraphIndices)
            Console.WriteLine(idx);
    }
}

// Callback that records the paragraph index of each match.
class CollectParagraphIndicesCallback : IReplacingCallback
{
    private readonly Document _document;
    private readonly List<int> _indices;

    public CollectParagraphIndicesCallback(Document document, List<int> indices)
    {
        _document = document;
        _indices = indices;
    }

    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
    {
        // Find the paragraph that contains the match.
        Paragraph paragraph = (Paragraph)args.MatchNode.GetAncestor(NodeType.Paragraph);

        if (paragraph != null)
        {
            // Get the index of the paragraph within the main body of the first section.
            int paragraphIndex = _document.FirstSection.Body.Paragraphs.IndexOf(paragraph);

            // Add the index if it hasn't been recorded yet.
            if (paragraphIndex >= 0 && !_indices.Contains(paragraphIndex))
                _indices.Add(paragraphIndex);
        }

        // Do not modify the document text.
        args.Replacement = args.Match.Value;
        return ReplaceAction.Replace;
    }
}
