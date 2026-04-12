using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a blank document and add sample markdown headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("# Title");
        builder.Writeln("Some introductory text.");
        builder.Writeln("## Section One");
        builder.Writeln("Content of section one.");
        builder.Writeln("### Subsection A");
        builder.Writeln("More details.");
        builder.Writeln("## Section Two");
        builder.Writeln("Final content.");

        // Regular expression to match markdown headings (levels 1‑6).
        // Group 1: the leading '#' characters.
        // Group 2: the heading text.
        string pattern = @"^(#{1,6})\s*(.+)$";
        Regex regex = new Regex(pattern, RegexOptions.Multiline);

        // Set up find‑replace options with a custom callback that applies the appropriate heading style.
        FindReplaceOptions options = new FindReplaceOptions(new MarkdownHeadingReplacer());

        // Perform the replacement. The callback supplies the replacement text and formatting.
        int replacementCount = doc.Range.Replace(regex, string.Empty, options);

        // Validate that at least one heading was converted.
        if (replacementCount == 0)
            throw new InvalidOperationException("No markdown headings were found to replace.");

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ConvertedHeadings.docx");
        doc.Save(outputPath);
    }

    // Callback that replaces a markdown heading with plain text and applies the corresponding Word heading style.
    private class MarkdownHeadingReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Extract the matched groups.
            Match match = args.Match;
            string hashes = match.Groups[1].Value;   // e.g., "##"
            string headingText = match.Groups[2].Value.Trim(); // e.g., "Section One"

            // Determine heading level from the number of '#' characters.
            int level = hashes.Length; // 1‑6

            // Set the replacement text (the heading without markdown symbols).
            args.Replacement = headingText;

            // Locate the paragraph that contains the match.
            Node current = args.MatchNode;
            while (current != null && current.NodeType != NodeType.Paragraph)
                current = current.ParentNode;

            if (current is Paragraph paragraph)
            {
                // Apply the appropriate built‑in heading style (Heading1 … Heading6).
                string styleName = $"Heading{level}";
                if (Enum.TryParse(styleName, out StyleIdentifier styleId))
                    paragraph.ParagraphFormat.StyleIdentifier = styleId;
            }

            return ReplaceAction.Replace;
        }
    }
}
