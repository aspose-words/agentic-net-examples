using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Define output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with headings and sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First heading (level 1).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1 - Section A");

        // Some body text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Paragraph under Heading 1.");

        // Insert a section break (new page) to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second heading (level 2) in the new section.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2 - Section B");

        // More body text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Paragraph under Heading 2.");

        // Insert another heading (level 3) without a new section.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 3 - Same Section");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Paragraph under Heading 3.");

        // Prepare HtmlSaveOptions to split by both headings and sections.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Combine the flags using bitwise OR.
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph | DocumentSplitCriteria.SectionBreak,
            // Split at heading levels 1 and 2 (adjust as needed).
            DocumentSplitHeadingLevel = 2
        };

        // Save the document; Aspose.Words will generate multiple HTML files.
        string baseFileName = Path.Combine(outputDir, "CombinedSplit.html");
        doc.Save(baseFileName, saveOptions);

        // Validate that split parts were created.
        // The main file plus at least one part file should exist.
        string[] htmlFiles = Directory.GetFiles(outputDir, "CombinedSplit*.html");
        if (htmlFiles.Length < 2)
        {
            throw new InvalidOperationException("Expected multiple split HTML files, but fewer were found.");
        }

        // Output the list of generated files (optional, for verification).
        Console.WriteLine("Generated split HTML files:");
        foreach (string file in htmlFiles.OrderBy(f => f))
        {
            Console.WriteLine(" - " + Path.GetFileName(file));
        }
    }
}
