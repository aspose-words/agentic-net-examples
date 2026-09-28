using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with headings and page breaks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is some text in chapter 1. It will span multiple lines to ensure content.");

        // Force a page break.
        builder.InsertBreak(BreakType.PageBreak);

        // Second heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 2 follows. More text to fill the page.");

        // Force another page break.
        builder.InsertBreak(BreakType.PageBreak);

        // Third heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 3");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Final chapter content.");

        // Prepare output folder.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Configure HTML save options to split on headings and page breaks.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        // Combine HeadingParagraph and PageBreak criteria so each part starts on a new page.
        saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph | DocumentSplitCriteria.PageBreak;
        // Export headers/footers per section (optional, kept from original example).
        saveOptions.ExportHeadersFootersMode = ExportHeadersFootersMode.PerSection;

        // Save the document; Aspose.Words will generate multiple HTML files.
        string mainFilePath = Path.Combine(outputDir, "Document.html");
        doc.Save(mainFilePath, saveOptions);

        // Validate that split files were created (main file + at least one split part).
        string[] generatedFiles = Directory.GetFiles(outputDir, "Document*.html");
        if (generatedFiles.Length < 2)
        {
            throw new InvalidOperationException("Expected split HTML files were not generated.");
        }

        // List generated files.
        foreach (var file in generatedFiles)
        {
            Console.WriteLine($"Generated: {file}");
        }
    }
}
