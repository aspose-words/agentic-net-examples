using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a blank document and add markdown‑style headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("# Document Title");
        builder.Writeln("Some introductory paragraph.");
        builder.Writeln("## First Section");
        builder.Writeln("Content of the first section.");
        builder.Writeln("### Subsection A");
        builder.Writeln("Details of subsection A.");
        builder.Writeln("## Second Section");
        builder.Writeln("Content of the second section.");

        // Regular expression that matches markdown headings (levels 1‑6).
        Regex headingRegex = new Regex(@"^(#{1,6})\s+(.*)$", RegexOptions.Multiline);

        // Configure find‑replace to use a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new HeadingReplacer()
        };

        // Perform the replacement. The callback will set the appropriate heading style.
        int replacedCount = doc.Range.Replace(headingRegex, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No markdown headings were found for replacement.");

        // Save the resulting document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }

    // Callback that converts a markdown heading to a Word heading style.
    private class HeadingReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Group 1 contains the series of '#' characters; its length determines the heading level (1‑6).
            string hashGroup = args.Match.Groups[1].Value;
            int level = hashGroup.Length;

            // Group 2 contains the actual heading text.
            string headingText = args.Match.Groups[2].Value.Trim();

            // Apply the corresponding built‑in heading style to the paragraph that holds the match.
            if (args.MatchNode?.ParentNode is Paragraph paragraph)
            {
                // Heading styles are sequential in the StyleIdentifier enum.
                StyleIdentifier styleId = (StyleIdentifier)((int)StyleIdentifier.Heading1 + level - 1);
                paragraph.ParagraphFormat.StyleIdentifier = styleId;
            }

            // Replace the markdown markup with plain heading text.
            args.Replacement = headingText;
            return ReplaceAction.Replace;
        }
    }
}
