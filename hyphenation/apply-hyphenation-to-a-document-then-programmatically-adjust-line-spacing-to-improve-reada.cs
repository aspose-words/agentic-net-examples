using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Output file path (relative to the executable's working directory)
        string outputPath = "HyphenatedLineSpacing.docx";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure a narrow page width to force line wrapping.
        builder.PageSetup.PageWidth = 200;   // 200 points (~2.78 inches)
        builder.PageSetup.LeftMargin = 20;   // 20 points margin
        builder.PageSetup.RightMargin = 20;  // 20 points margin

        // Set a readable font size.
        builder.Font.Size = 12;

        // Configure line spacing to improve readability.
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 18; // 1.5 lines (default line height is 12 points)

        // Write a long paragraph containing words that can be hyphenated.
        builder.Writeln(
            "This is a sample paragraph containing a series of long words such as " +
            "antidisestablishmentarianism, pneumonoultramicroscopicsilicovolcanoconiosis, " +
            "and supercalifragilisticexpialidocious to demonstrate hyphenation.");

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2; // Allow up to 2 consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenationZone = 360;      // 0.25 inch from the right margin (default value).
        doc.HyphenationOptions.HyphenateCaps = true;     // Hyphenate words in all caps if present.

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
