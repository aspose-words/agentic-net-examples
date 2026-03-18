using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class ReplaceAfterHeading
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Find the heading after which replacements should be applied.
        // In this example we look for a paragraph whose text exactly matches "My Heading".
        Paragraph? headingParagraph = null;
        NodeCollection allParagraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in allParagraphs)
        {
            if (para.GetText().Trim() == "My Heading")
            {
                headingParagraph = para;
                break;
            }
        }

        if (headingParagraph == null)
            throw new InvalidOperationException("Heading not found.");

        // Determine the index of the heading paragraph within the document order.
        int headingIndex = allParagraphs.IndexOf(headingParagraph);

        // Set up a callback that will replace only matches occurring after the heading.
        var callback = new ReplaceAfterHeadingCallback(headingIndex, allParagraphs);
        FindReplaceOptions options = new FindReplaceOptions(callback);

        // Example: replace every occurrence of the word "foo" with "bar" after the heading.
        doc.Range.Replace(new Regex(@"\bfoo\b", RegexOptions.IgnoreCase), "bar", options);

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Callback implementation that checks the position of each match.
    private class ReplaceAfterHeadingCallback : IReplacingCallback
    {
        private readonly int _headingIndex;
        private readonly NodeCollection _paragraphs;

        public ReplaceAfterHeadingCallback(int headingIndex, NodeCollection paragraphs)
        {
            _headingIndex = headingIndex;
            _paragraphs = paragraphs;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Find the paragraph that contains the current match.
            Paragraph matchParagraph = (Paragraph)args.MatchNode.GetAncestor(NodeType.Paragraph);

            // Determine its index in the same collection used for the heading.
            int matchIndex = _paragraphs.IndexOf(matchParagraph);

            // Perform replacement only if the match is after the heading.
            if (matchIndex > _headingIndex)
            {
                // Optionally modify the replacement text here.
                // args.Replacement = $"[Replaced] {args.Replacement}";
                return ReplaceAction.Replace;
            }

            // Skip replacement for matches before or within the heading.
            return ReplaceAction.Skip;
        }
    }
}
