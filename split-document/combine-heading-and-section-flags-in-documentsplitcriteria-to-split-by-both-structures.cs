using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with headings and explicit section breaks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Heading level 1 – will be a split point.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        // Normal paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Paragraph under heading 1.");

        // Insert a section break – will also be a split point.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Heading level 2 – will be a split point (up to level 2).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Paragraph under heading 2.");

        // Another section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Heading level 3 – will NOT be a split point because we limit to level 2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 3");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Paragraph under heading 3.");

        // Save the original document (optional, just for reference).
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        doc.Save(sourcePath);

        // Configure HtmlSaveOptions to split by both headings and sections.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph | DocumentSplitCriteria.SectionBreak,
            DocumentSplitHeadingLevel = 2 // Split at Heading 1 and Heading 2.
        };

        // Save the document; Aspose.Words will create multiple HTML files.
        string baseHtmlPath = Path.Combine(outputDir, "SplitDocument.html");
        doc.Save(baseHtmlPath, saveOptions);

        // Validate that split files were created.
        string[] splitFiles = Directory.GetFiles(outputDir, "SplitDocument*.html");
        if (splitFiles.Length < 2)
            throw new InvalidOperationException("Expected multiple split HTML files, but none were created.");

        // Output the list of generated files (for debugging purposes).
        Console.WriteLine("Generated split files:");
        foreach (string file in splitFiles)
            Console.WriteLine(file);
    }
}
