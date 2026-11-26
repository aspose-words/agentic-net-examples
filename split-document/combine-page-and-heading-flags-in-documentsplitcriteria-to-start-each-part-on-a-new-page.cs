using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with headings and a page break.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First heading and paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Paragraph under heading 1.");

        // Insert a page break.
        builder.InsertBreak(BreakType.PageBreak);

        // Second heading and paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Paragraph under heading 2.");

        // Save the source document.
        string sourcePath = Path.Combine(outputDir, "Sample.docx");
        doc.Save(sourcePath);

        // Set up HTML save options with combined split criteria (page break + heading paragraph).
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.PageBreak | DocumentSplitCriteria.HeadingParagraph;
        saveOptions.DocumentSplitHeadingLevel = 1; // Split at heading level 1.
        saveOptions.ExportPageMargins = true;
        saveOptions.ExportImagesAsBase64 = true;

        // Save the document, which will generate split HTML files.
        string baseHtmlPath = Path.Combine(outputDir, "SplitDocument.html");
        doc.Save(baseHtmlPath, saveOptions);

        // Validate that split files were created (at least two parts expected).
        string[] splitFiles = Directory.GetFiles(outputDir, "SplitDocument*.html");
        if (splitFiles.Length < 2)
            throw new Exception("Expected split HTML files were not created.");

        // Output the count of generated files.
        Console.WriteLine($"Split operation completed. Generated {splitFiles.Length} HTML files in '{outputDir}'.");
    }
}
