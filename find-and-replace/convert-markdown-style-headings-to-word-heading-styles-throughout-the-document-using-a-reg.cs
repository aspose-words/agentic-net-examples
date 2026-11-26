using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing markdown style headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("# Heading 1");
        builder.Writeln("Paragraph under heading 1.");
        builder.Writeln("## Heading 2");
        builder.Writeln("Paragraph under heading 2.");
        builder.Writeln("### Heading 3");
        builder.Writeln("Paragraph under heading 3.");
        builder.Writeln("Normal paragraph without heading.");

        // Regular expression to match markdown headings.
        // Captures the leading hashes (1-6) and the heading text.
        Regex headingRegex = new Regex(@"^(?<hashes>#{1,6})\s+(?<text>.+)$",
                                        RegexOptions.Multiline);

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new HeadingReplacingCallback();

        // Perform the replacement. The callback will replace the text and apply the style.
        int replacedCount = doc.Range.Replace(headingRegex, string.Empty, options);

        // Validate that at least one heading was converted.
        if (replacedCount == 0)
            throw new InvalidOperationException("No markdown headings were found to replace.");

        // Save the resulting document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }

    // Callback that replaces markdown heading syntax with plain text and applies Word heading styles.
    private class HeadingReplacingCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Extract captured groups from the regex match.
            Match match = args.Match;
            string hashes = match.Groups["hashes"].Value;
            string headingText = match.Groups["text"].Value;

            // Determine heading level based on the number of leading '#'.
            int level = hashes.Length; // 1‑6

            // Set the replacement text to the heading text without markdown symbols.
            args.Replacement = headingText;

            // Locate the paragraph that contains the matched run.
            Node? matchNode = args.MatchNode;
            if (matchNode != null)
            {
                Paragraph? paragraph = matchNode.GetAncestor(NodeType.Paragraph) as Paragraph;
                if (paragraph != null)
                {
                    // Apply the corresponding built‑in Word heading style (e.g., "Heading 1").
                    paragraph.ParagraphFormat.StyleName = $"Heading {level}";
                }
            }

            // Instruct Aspose.Words to replace the match with the new text.
            return ReplaceAction.Replace;
        }
    }
}
