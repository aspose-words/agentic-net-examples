using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace MarkdownHeadingConverter
{
    // Callback that replaces markdown headings with plain text and applies Word heading styles.
    public class HeadingReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The match contains the whole markdown heading line.
            Match match = args.Match;
            // Group 1 = sequence of '#' characters, Group 2 = heading text.
            string hashes = match.Groups[1].Value;
            string headingText = match.Groups[2].Value;

            // Determine heading level (1‑6) based on number of '#'.
            int level = hashes.Length;
            if (level < 1 || level > 6)
                level = 1; // Fallback to Heading1 if unexpected.

            // Replace the markdown markup with just the heading text.
            args.Replacement = headingText;

            // Find the paragraph that contains the match and set its style.
            Node? matchNode = args.MatchNode;
            if (matchNode != null)
            {
                Paragraph? paragraph = matchNode.GetAncestor(NodeType.Paragraph) as Paragraph;
                if (paragraph != null)
                {
                    // Map level to the corresponding built‑in heading style.
                    paragraph.ParagraphFormat.StyleIdentifier = level switch
                    {
                        1 => StyleIdentifier.Heading1,
                        2 => StyleIdentifier.Heading2,
                        3 => StyleIdentifier.Heading3,
                        4 => StyleIdentifier.Heading4,
                        5 => StyleIdentifier.Heading5,
                        6 => StyleIdentifier.Heading6,
                        _ => StyleIdentifier.Heading1
                    };
                }
            }

            return ReplaceAction.Replace;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a sample document containing markdown style headings.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("# First Level Heading");
            builder.Writeln("Paragraph under first level heading.");
            builder.Writeln("## Second Level Heading");
            builder.Writeln("Paragraph under second level heading.");
            builder.Writeln("### Third Level Heading");
            builder.Writeln("Another paragraph.");

            // Save the original sample (optional, just for demonstration).
            string inputPath = "markdown_input.docx";
            doc.Save(inputPath);

            // Regular expression to match markdown headings (lines starting with 1‑6 '#').
            Regex headingRegex = new Regex(@"^(#{1,6})\s+(.*)$", RegexOptions.Multiline);

            // Set up find‑replace options with the custom callback.
            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new HeadingReplacer()
            };

            // Perform the replacement. The replacement string is ignored because the callback supplies it.
            int replacedCount = doc.Range.Replace(headingRegex, string.Empty, options);

            // Validate that at least one heading was processed.
            if (replacedCount == 0)
                throw new InvalidOperationException("No markdown headings were found for replacement.");

            // Save the transformed document.
            string outputPath = "converted_headings.docx";
            doc.Save(outputPath);
        }
    }
}
