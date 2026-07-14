using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new document and build content with headings and sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First heading (Heading 1) and some text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1 - Start of Document");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Paragraph under Heading 1.");

        // Insert a section break (new page) and add another heading.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2 - New Section");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Paragraph under Heading 2 in a new section.");

        // Add a third heading without a section break.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 3 - Same Section");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Paragraph under Heading 3.");

        // Set up HtmlSaveOptions to split by both headings and section breaks.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Combine flags using bitwise OR.
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph | DocumentSplitCriteria.SectionBreak,
            // Split at heading levels 1 and 2 (adjust as needed).
            DocumentSplitHeadingLevel = 2
        };

        // Base file name for the split output.
        string baseFileName = Path.Combine(outputDir, "CombinedSplit.html");

        // Save the document; Aspose.Words will create multiple HTML files.
        doc.Save(baseFileName, saveOptions);

        // Verify that split files were created.
        string[] splitFiles = Directory.GetFiles(outputDir, "CombinedSplit*.html");
        if (splitFiles.Length < 2)
            throw new InvalidOperationException("Expected multiple split HTML files, but fewer were created.");

        // Output the names of the generated files.
        Console.WriteLine("Split HTML files created:");
        foreach (string file in splitFiles.OrderBy(f => f))
        {
            Console.WriteLine(Path.GetFileName(file));
        }
    }
}
