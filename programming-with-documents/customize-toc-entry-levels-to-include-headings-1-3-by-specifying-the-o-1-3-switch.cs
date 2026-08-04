using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents that includes heading levels 1 through 3.
        // \\o "1-3" specifies the range of heading levels.
        // \\h makes entries hyperlinks, \\z hides page numbers in web layout, \\u builds the TOC from outline levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add sample headings with styles Heading 1, Heading 2, and Heading 3.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 1.1");
        builder.Writeln("Heading 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 1.2.1");
        builder.Writeln("Heading 1.2.2");

        // Update all fields in the document so the TOC reflects the added headings.
        doc.UpdateFields();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(outputDir);

        // Save the document to the output folder.
        string outputPath = Path.Combine(outputDir, "CustomToc.docx");
        doc.Save(outputPath);
    }
}
