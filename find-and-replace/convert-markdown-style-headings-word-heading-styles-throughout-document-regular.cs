using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Replacing;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load a Markdown document. The MarkdownLoadOptions enable proper handling of Markdown syntax.
        Document doc = new Document("input.md", new MarkdownLoadOptions());

        // Regular expression that matches Markdown headings:
        //   ^(#{1,6})\s*(.+)$
        // Group 1 – the sequence of '#' characters (determines the heading level).
        // Group 2 – the heading text.
        Regex headingRegex = new Regex(@"^(#{1,6})\s*(.+)$", RegexOptions.Multiline);

        // Configure find/replace to use a callback that will set the appropriate Word heading style.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new HeadingReplacingCallback();

        // Perform the replace. The replacement string "$2" keeps only the heading text.
        doc.Range.Replace(headingRegex, "$2", options);

        // Save the result as a Word document.
        doc.Save("output.docx");
    }

    // Callback that runs for each regex match.
    private class HeadingReplacingCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Determine heading level from the length of the '#' sequence (group 1).
            string hashes = args.Match.Groups[1].Value;
            int level = hashes.Length; // Valid values: 1‑6

            // The match node is a Run; its parent is the Paragraph that we need to style.
            Paragraph paragraph = (Paragraph)args.MatchNode.ParentNode;

            // Apply the corresponding built‑in heading style.
            switch (level)
            {
                case 1: paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1; break;
                case 2: paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2; break;
                case 3: paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3; break;
                case 4: paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading4; break;
                case 5: paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading5; break;
                case 6: paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading6; break;
            }

            // Replace the whole match with the captured heading text (group 2).
            args.Replacement = args.Match.Groups[2].Value;

            return ReplaceAction.Replace;
        }
    }
}
