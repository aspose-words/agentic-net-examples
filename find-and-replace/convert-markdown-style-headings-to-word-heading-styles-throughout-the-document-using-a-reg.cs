using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;

namespace MarkdownHeadingConverter
{
    // Callback that converts a markdown heading to a Word heading style.
    public class HeadingReplacingCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Ensure the match node exists.
            if (args.MatchNode == null)
                return ReplaceAction.Skip;

            // Find the paragraph that contains the match.
            Paragraph paragraph = args.MatchNode.GetAncestor(NodeType.Paragraph) as Paragraph;
            if (paragraph == null)
                return ReplaceAction.Skip;

            // Extract heading level (number of '#') and heading text.
            Match match = args.Match;
            int level = match.Groups[1].Value.Length; // 1‑6
            string headingText = match.Groups[2].Value.Trim();

            // Apply the appropriate heading style.
            paragraph.ParagraphFormat.StyleIdentifier = level switch
            {
                1 => StyleIdentifier.Heading1,
                2 => StyleIdentifier.Heading2,
                3 => StyleIdentifier.Heading3,
                4 => StyleIdentifier.Heading4,
                5 => StyleIdentifier.Heading5,
                6 => StyleIdentifier.Heading6,
                _ => StyleIdentifier.Normal
            };

            // Replace the matched markdown text with plain heading text.
            args.Replacement = headingText;

            // Indicate that the default replacement should be performed.
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
            builder.Writeln("# Document Title");
            builder.Writeln("Some introductory paragraph.");
            builder.Writeln("## Chapter One");
            builder.Writeln("Content of chapter one.");
            builder.Writeln("### Section 1.1");
            builder.Writeln("Details of section 1.1.");
            builder.Writeln("## Chapter Two");
            builder.Writeln("Content of chapter two.");

            // Regex that matches markdown headings (levels 1‑6).
            Regex headingRegex = new Regex(@"^(#{1,6})\s*(.+)$", RegexOptions.Multiline);

            // Set up find‑replace options with the custom callback.
            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new HeadingReplacingCallback()
            };

            // Perform the replacement. The callback supplies the replacement text.
            int replacedCount = doc.Range.Replace(headingRegex, string.Empty, options);

            // Validate that at least one heading was converted.
            if (replacedCount == 0)
                throw new InvalidOperationException("No markdown headings were found to replace.");

            // Save the resulting document.
            const string outputPath = "ConvertedHeadings.docx";
            doc.Save(outputPath, SaveFormat.Docx);
        }
    }
}
