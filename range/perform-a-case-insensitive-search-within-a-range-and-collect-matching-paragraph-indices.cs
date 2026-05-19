using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace AsposeWordsRangeSearch
{
    // Callback that records the index of each paragraph containing a match.
    class ParagraphMatchCollector : IReplacingCallback
    {
        private readonly Document _document;
        private readonly List<int> _paragraphIndices;

        public ParagraphMatchCollector(Document document, List<int> paragraphIndices)
        {
            _document = document;
            _paragraphIndices = paragraphIndices;
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Find the paragraph that contains the match.
            Paragraph paragraph = (Paragraph)args.MatchNode.GetAncestor(NodeType.Paragraph);
            if (paragraph != null)
            {
                // Determine the paragraph's index within the body.
                int index = _document.FirstSection.Body.Paragraphs.IndexOf(paragraph);
                if (index >= 0 && !_paragraphIndices.Contains(index))
                    _paragraphIndices.Add(index);
            }

            // Skip actual replacement – we only want to collect matches.
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
            builder.Writeln("Aspose.Words provides powerful document processing.");
            builder.Writeln("Another line without the keyword.");
            builder.Writeln("The word aspose appears here in lower case.");
            builder.Writeln("Final paragraph with ASPOSE in upper case.");

            // Prepare a list to hold matching paragraph indices.
            List<int> matchingParagraphIndices = new List<int>();

            // Set up find/replace options for a case‑insensitive search.
            FindReplaceOptions options = new FindReplaceOptions
            {
                MatchCase = false, // case‑insensitive
                ReplacingCallback = new ParagraphMatchCollector(doc, matchingParagraphIndices)
            };

            // Perform a replace operation where the replacement text is identical to the search text.
            // This triggers the callback without altering the document.
            doc.Range.Replace("aspose", "aspose", options);

            // Output the collected paragraph indices.
            Console.WriteLine("Paragraph indices containing the term \"aspose\" (case‑insensitive):");
            foreach (int index in matchingParagraphIndices)
            {
                Console.WriteLine(index);
            }
        }
    }
}
