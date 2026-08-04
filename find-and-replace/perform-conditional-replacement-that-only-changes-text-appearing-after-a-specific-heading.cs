using System;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and populate it with sample content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Text before the target heading – should NOT be replaced.
        builder.Writeln("Before heading placeholder: PLACEHOLDER");
        builder.Writeln("Another before: PLACEHOLDER");

        // Insert the target heading.
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Target Heading");
        // Return to normal style for following paragraphs.
        builder.Font.Size = 12;
        builder.Font.Bold = false;
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // Text after the target heading – should be replaced.
        builder.Writeln("After heading placeholder: PLACEHOLDER");
        builder.Writeln("More text with PLACEHOLDER inside.");

        // Set up the find‑replace options with a custom callback.
        var options = new FindReplaceOptions
        {
            ReplacingCallback = new ConditionalReplacer("Target Heading")
        };

        // Perform the replacement.
        int replacedCount = doc.Range.Replace("PLACEHOLDER", "REPLACED", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement after the heading.");

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Callback that replaces only when the match occurs after a specific heading.
    private class ConditionalReplacer : IReplacingCallback
    {
        private readonly string _headingText;

        public ConditionalReplacer(string headingText)
        {
            _headingText = headingText;
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Locate the paragraph that contains the start of the match.
            Node matchNode = args.MatchNode;
            while (matchNode != null && matchNode.NodeType != NodeType.Paragraph)
                matchNode = matchNode.ParentNode;

            if (matchNode == null)
                return ReplaceAction.Skip; // Safety check.

            // Walk backwards through preceding siblings to find the heading.
            Node current = matchNode.PreviousSibling;
            while (current != null)
            {
                if (current.NodeType == NodeType.Paragraph)
                {
                    Paragraph para = (Paragraph)current;
                    if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1 &&
                        para.GetText().Trim().Equals(_headingText, StringComparison.Ordinal))
                    {
                        // Heading found before the match – perform replacement.
                        args.Replacement = "REPLACED";
                        return ReplaceAction.Replace;
                    }
                }
                current = current.PreviousSibling;
            }

            // No preceding heading found – skip this match.
            return ReplaceAction.Skip;
        }
    }
}
