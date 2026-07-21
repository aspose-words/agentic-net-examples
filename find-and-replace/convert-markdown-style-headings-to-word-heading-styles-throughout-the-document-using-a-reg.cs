using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Drawing; // Required by the task specification
using Newtonsoft.Json; // Required by the task specification

namespace MarkdownHeadingConverter
{
    // Callback that converts markdown headings to Word heading styles.
    public class HeadingReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Determine heading level by counting leading '#'.
            string matchText = args.Match.Value;
            int level = 0;
            foreach (char c in matchText)
            {
                if (c == '#')
                    level++;
                else
                    break;
            }

            // Fallback to normal replace if no heading marker found.
            if (level == 0)
                return ReplaceAction.Skip;

            // Extract the heading title without markdown symbols.
            string headingText = matchText.Substring(level).Trim();

            // Set the replacement text (plain heading title).
            args.Replacement = headingText;

            // Apply the appropriate Word heading style to the paragraph containing the match.
            var run = args.MatchNode as Run;
            if (run != null && run.ParentNode is Paragraph paragraph)
            {
                switch (level)
                {
                    case 1:
                        paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
                        break;
                    case 2:
                        paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
                        break;
                    case 3:
                        paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
                        break;
                    case 4:
                        paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading4;
                        break;
                    case 5:
                        paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading5;
                        break;
                    case 6:
                        paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading6;
                        break;
                    default:
                        // For levels beyond 6, use Heading6 as the highest available style.
                        paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading6;
                        break;
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

            builder.Writeln("# Introduction");
            builder.Writeln("This is an introductory paragraph.");
            builder.Writeln("## Subsection A");
            builder.Writeln("Details about subsection A.");
            builder.Writeln("### Sub-subsection A1");
            builder.Writeln("More details.");
            builder.Writeln("## Subsection B");
            builder.Writeln("Details about subsection B.");
            builder.Writeln("# Conclusion");
            builder.Writeln("Final remarks.");

            // Save the original document (optional, for reference).
            doc.Save("OriginalMarkdown.docx");

            // Prepare regex to match markdown headings (lines starting with 1-6 '#').
            Regex headingRegex = new Regex(@"^(#{1,6})\s+.*$", RegexOptions.Multiline);

            // Set up find-replace options with our custom callback.
            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new HeadingReplacer()
            };

            // Perform the replacement.
            int replacedCount = doc.Range.Replace(headingRegex, "", options);

            // Validate that replacements occurred.
            if (replacedCount == 0)
                throw new InvalidOperationException("No markdown headings were found for replacement.");

            // Save the transformed document.
            doc.Save("ConvertedHeadings.docx");
        }
    }
}
