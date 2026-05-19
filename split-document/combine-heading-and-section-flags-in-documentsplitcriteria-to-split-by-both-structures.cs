using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample document with headings and sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section with headings.
        builder.Writeln("Content before first heading.");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1 - Section 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2 - Section 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content after headings in first section.");

        // Insert a section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section with headings.
        builder.Writeln("Content before second heading.");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1 - Section 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2 - Section 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content after headings in second section.");

        // Configure HtmlSaveOptions to split by both headings and sections.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph | DocumentSplitCriteria.SectionBreak,
            DocumentSplitHeadingLevel = 2 // Split at Heading 1 and Heading 2 levels.
        };

        // Save the document; Aspose.Words will create multiple HTML files.
        string baseFileName = Path.Combine(artifactsDir, "CombinedSplit.html");
        doc.Save(baseFileName, saveOptions);

        // Verify that split files were created.
        string[] splitFiles = Directory.GetFiles(artifactsDir, "CombinedSplit*.html");
        if (splitFiles.Length < 2)
            throw new Exception("Expected multiple split HTML files, but only one was found.");

        // Output the names of the generated files.
        Console.WriteLine("Generated split HTML files:");
        foreach (string file in splitFiles.OrderBy(f => f))
        {
            Console.WriteLine(Path.GetFileName(file));
        }
    }
}
