using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a heading that will act as the marker.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Target Heading");

        // Paragraphs after the target heading that contain the text to be replaced.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a PLACEHOLDER that should be replaced.");
        builder.Writeln("Another PLACEHOLDER appears here.");

        // Insert a different heading – text after this should NOT be changed.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Other Heading");

        // Paragraphs after the other heading – should remain untouched.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This PLACEHOLDER must stay as is.");

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new ConditionalReplacer("Target Heading")
        };

        // Perform the replacement only on matches that satisfy the callback condition.
        int replacedCount = doc.Range.Replace(new Regex(@"PLACEHOLDER"), "REPLACED", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No replacements were made; the conditional logic may be incorrect.");

        // Save the modified document to the local folder.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);
    }

    // Callback that replaces a match only if it appears after a specific heading.
    private class ConditionalReplacer : IReplacingCallback
    {
        private readonly string _headingText;

        public ConditionalReplacer(string headingText)
        {
            _headingText = headingText ?? throw new ArgumentNullException(nameof(headingText));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Find the paragraph that contains the current match.
            Paragraph matchParagraph = args.MatchNode.GetAncestor(typeof(Paragraph)) as Paragraph;
            if (matchParagraph == null)
                return ReplaceAction.Skip;

            // Walk backwards through preceding sibling nodes to locate the heading.
            Node? current = matchParagraph.PreviousSibling;
            while (current != null)
            {
                if (current.NodeType == NodeType.Paragraph)
                {
                    Paragraph para = (Paragraph)current;
                    // Check if this paragraph is a heading with the required text.
                    if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1 &&
                        para.GetText().Trim().Equals(_headingText, StringComparison.Ordinal))
                    {
                        // The heading was found before this match – perform replacement.
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
