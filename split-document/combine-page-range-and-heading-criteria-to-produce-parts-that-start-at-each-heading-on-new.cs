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

        // -----------------------------------------------------------------
        // 1. Create a sample source document with headings that start on new pages.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Chapter 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 1.");

        // Force a page break before the next heading.
        builder.InsertBreak(BreakType.PageBreak);

        // Chapter 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 2.");

        // Another page break.
        builder.InsertBreak(BreakType.PageBreak);

        // Chapter 3
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 3");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 3.");

        // Save the source document.
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Configure HtmlSaveOptions to split at headings and page breaks.
        // -----------------------------------------------------------------
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Combine heading and page‑break criteria.
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph | DocumentSplitCriteria.PageBreak,
            // Split at Heading 1 paragraphs.
            DocumentSplitHeadingLevel = 1
        };

        // Save the document; Aspose.Words will create separate HTML files:
        // Source.html, Source-01.html, Source-02.html, …
        string htmlBasePath = Path.Combine(outputDir, "Source.html");
        sourceDoc.Save(htmlBasePath, saveOptions);

        // -----------------------------------------------------------------
        // 3. Validate that split parts were created.
        // -----------------------------------------------------------------
        string[] splitFiles = Directory.GetFiles(outputDir, "Source-*.html");
        Console.WriteLine($"Split parts created: {splitFiles.Length}");
        foreach (string file in splitFiles)
        {
            Console.WriteLine($" - {Path.GetFileName(file)}");
        }
    }
}
