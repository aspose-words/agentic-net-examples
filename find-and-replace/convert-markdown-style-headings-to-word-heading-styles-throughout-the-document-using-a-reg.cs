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
        builder.Writeln("# Title");
        builder.Writeln("This is a normal paragraph.");
        builder.Writeln("## Section");
        builder.Writeln("Content under the section.");
        builder.Writeln("### Subsection");
        builder.Writeln("More detailed content.");

        // Regular expression to match markdown headings (levels 1‑6).
        Regex headingRegex = new Regex(@"^(#{1,6})\s*(.+)$", RegexOptions.Multiline);

        // Configure replace options with a callback that applies Word heading styles.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new HeadingCallback()
        };

        // Perform the find‑and‑replace operation.
        int replacedCount = doc.Range.Replace(headingRegex, string.Empty, options);
        if (replacedCount == 0)
            throw new InvalidOperationException("No markdown headings were found for replacement.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        doc.Save(outputPath);
    }
}

// Callback that converts a markdown heading to a Word heading style.
public class HeadingCallback : IReplacingCallback
{
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Determine heading level from the number of leading '#'.
        string hashGroup = args.Match.Groups[1].Value;
        int level = hashGroup.Length; // 1 to 6

        // Extract the heading text without markdown symbols.
        string headingText = args.Match.Groups[2].Value.Trim();

        // Replace the markdown syntax with plain heading text.
        args.Replacement = headingText;

        // Apply the corresponding Word heading style to the paragraph.
        if (args.MatchNode?.ParentNode is Paragraph paragraph)
        {
            StyleIdentifier styleId = level switch
            {
                1 => StyleIdentifier.Heading1,
                2 => StyleIdentifier.Heading2,
                3 => StyleIdentifier.Heading3,
                4 => StyleIdentifier.Heading4,
                5 => StyleIdentifier.Heading5,
                6 => StyleIdentifier.Heading6,
                _ => StyleIdentifier.Normal
            };
            paragraph.ParagraphFormat.StyleIdentifier = styleId;
        }

        return ReplaceAction.Replace;
    }
}
