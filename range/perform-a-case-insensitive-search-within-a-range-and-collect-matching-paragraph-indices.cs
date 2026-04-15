using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace AsposeWordsRangeSearch
{
    // Callback that records the indices of paragraphs containing a match.
    class MatchCollector : IReplacingCallback
    {
        private readonly Document _doc;
        private readonly List<int> _indices;

        public MatchCollector(Document doc, List<int> indices)
        {
            _doc = doc;
            _indices = indices;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Find the paragraph that contains the match.
            Node current = args.MatchNode;
            while (current != null && current.NodeType != NodeType.Paragraph)
                current = current.ParentNode;

            if (current is Paragraph paragraph)
            {
                // Determine the paragraph's index within the main body.
                int index = _doc.FirstSection.Body.Paragraphs.IndexOf(paragraph);
                if (index >= 0 && !_indices.Contains(index))
                    _indices.Add(index);
            }

            // Skip actual replacement; we only want to collect matches.
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

            builder.Writeln("This is an Aspose example.");          // index 0
            builder.Writeln("Another example with aspose.");        // index 1
            builder.Writeln("No match here.");                     // index 2
            builder.Writeln("ASPOSE appears again.");              // index 3
            builder.Writeln("Final paragraph without keyword.");   // index 4

            // List to hold the indices of paragraphs that contain the search term.
            List<int> matchingParagraphIndices = new List<int>();

            // Set up find/replace options for a case‑insensitive search.
            FindReplaceOptions options = new FindReplaceOptions
            {
                MatchCase = false, // case‑insensitive
                ReplacingCallback = new MatchCollector(doc, matchingParagraphIndices)
            };

            // Perform a find operation. The replacement text is identical to the pattern,
            // and the callback skips the actual replacement.
            doc.Range.Replace("Aspose", "Aspose", options);

            // Output the collected paragraph indices.
            Console.WriteLine("Paragraph indices containing the word \"Aspose\" (case‑insensitive):");
            foreach (int idx in matchingParagraphIndices)
                Console.WriteLine(idx);

            // Optionally save the document to verify the content.
            doc.Save("SampleDocument.docx");
        }
    }
}
