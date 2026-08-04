using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("This is the first paragraph containing Aspose.");
        builder.Writeln("Second paragraph without the keyword.");
        builder.Writeln("Third paragraph with the word aspose in lower case.");
        builder.Writeln("Fourth paragraph with ASPose in upper case.");
        builder.Writeln("Fifth paragraph with no match.");

        // Prepare a collection to store unique paragraph indices where matches are found.
        HashSet<int> matchingIndices = new HashSet<int>();

        // Set up find/replace options for a case‑insensitive search.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false, // Ignore case.
            ReplacingCallback = new MatchRecorder(matchingIndices, doc)
        };

        // Perform a replace where the replacement text is identical to the pattern.
        // The callback will record matches and skip actual replacement.
        doc.Range.Replace("Aspose", "Aspose", options);

        // Output the collected paragraph indices.
        Console.WriteLine("Paragraph indices containing the term \"Aspose\" (case‑insensitive):");
        foreach (int index in matchingIndices)
        {
            Console.WriteLine(index);
        }

        // Save the (unchanged) document to demonstrate the lifecycle rule.
        doc.Save("Output.docx");
    }

    // Callback that records the index of the paragraph containing each match.
    private class MatchRecorder : IReplacingCallback
    {
        private readonly HashSet<int> _indices;
        private readonly Document _document;

        public MatchRecorder(HashSet<int> indices, Document document)
        {
            _indices = indices;
            _document = document;
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Find the paragraph that contains the match.
            Paragraph paragraph = (Paragraph)args.MatchNode.GetAncestor(NodeType.Paragraph);
            if (paragraph != null)
            {
                // Determine the paragraph's index within the body.
                int index = _document.FirstSection.Body.Paragraphs.IndexOf(paragraph);
                if (index >= 0)
                {
                    _indices.Add(index);
                }
            }

            // Skip replacement to keep the document unchanged.
            return ReplaceAction.Skip;
        }
    }
}
