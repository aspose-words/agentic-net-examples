using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace MarkdownHeadingConverter
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample document containing markdown style headings.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("# Title");
            builder.Writeln("Some introductory text.");
            builder.Writeln("## Section 1");
            builder.Writeln("Content of section 1.");
            builder.Writeln("### Subsection 1.1");
            builder.Writeln("More details.");
            builder.Writeln("## Section 2");
            builder.Writeln("Content of section 2.");

            const string inputPath = "input.docx";
            doc.Save(inputPath);

            // Load the document for processing.
            var loaded = new Document(inputPath);

            // Regex to match markdown headings (levels 1‑6).
            var headingRegex = new Regex(@"^(#{1,6})\s*(.+)$", RegexOptions.Multiline);

            var replaceOptions = new FindReplaceOptions
            {
                ReplacingCallback = new HeadingReplacingCallback()
            };

            // Perform the replace; the callback supplies the replacement text and styling.
            int replacedCount = loaded.Range.Replace(headingRegex, string.Empty, replaceOptions);

            if (replacedCount == 0)
                throw new InvalidOperationException("No markdown headings were found to replace.");

            const string outputPath = "output.docx";
            loaded.Save(outputPath);
        }
    }

    // Callback that converts a markdown heading to a Word heading style.
    internal class HeadingReplacingCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Determine heading level from the number of leading '#'.
            var match = args.Match;
            int level = match.Groups[1].Value.Length;
            string headingText = match.Groups[2].Value.Trim();

            // Replace the markdown syntax with plain heading text.
            args.Replacement = headingText;

            // Apply the appropriate Word heading style to the paragraph containing the match.
            if (args.MatchNode?.ParentNode is Paragraph paragraph)
            {
                switch (level)
                {
                    case 1: paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1; break;
                    case 2: paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2; break;
                    case 3: paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3; break;
                    case 4: paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading4; break;
                    case 5: paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading5; break;
                    case 6: paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading6; break;
                }
            }

            return ReplaceAction.Replace;
        }
    }
}
