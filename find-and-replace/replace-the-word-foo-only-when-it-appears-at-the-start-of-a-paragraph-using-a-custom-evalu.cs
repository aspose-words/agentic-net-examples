using System;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace FindReplaceStartOfParagraph
{
    // Callback that replaces only when the match is at the start of a paragraph.
    public class StartOfParagraphCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // We are only interested in matches that are inside a Run node.
            if (args.MatchNode.NodeType != NodeType.Run)
                return ReplaceAction.Skip;

            // Get the paragraph that contains the match.
            Paragraph paragraph = (Paragraph)args.MatchNode.GetAncestor(NodeType.Paragraph);
            if (paragraph == null)
                return ReplaceAction.Skip;

            // Get the paragraph text without the paragraph mark (character 0x07).
            string paragraphText = paragraph.GetText();
            if (paragraphText.EndsWith("\a"))
                paragraphText = paragraphText.Substring(0, paragraphText.Length - 1);

            // If the paragraph starts with the exact match value, allow replacement.
            if (paragraphText.StartsWith(args.Match.Value, StringComparison.Ordinal))
                return ReplaceAction.Replace;

            return ReplaceAction.Skip;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a sample document in memory.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("foo is at the start of this paragraph.");
            builder.Writeln("This paragraph contains foo inside.");
            builder.Writeln("foo appears again at the beginning.");
            builder.Writeln("No occurrence here.");
            builder.Writeln("foo"); // Paragraph that contains only the word.

            // Save the original document for inspection (optional).
            doc.Save("input.docx");

            // Set up find‑replace options with the custom callback.
            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new StartOfParagraphCallback()
            };

            // Replace "foo" with "bar" only when it appears at the start of a paragraph.
            int replacedCount = doc.Range.Replace("foo", "bar", options);

            // Ensure that at least one replacement was performed.
            if (replacedCount == 0)
                throw new InvalidOperationException("Expected at least one replacement, but none were made.");

            // Save the modified document.
            doc.Save("output.docx");
        }
    }
}
