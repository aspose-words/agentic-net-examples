using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a folder for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Initialize a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents that includes heading levels 1‑3.
        // \\o "1-3" specifies the range of heading levels to include.
        // \\h makes entries hyperlinks, \\z hides page numbers in web layout, \\u builds the TOC from outline levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add sample headings that will appear in the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 1.1");
        builder.Writeln("Heading 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 1.2.1");
        builder.Writeln("Heading 1.2.2");

        // Update fields so the TOC reflects the inserted headings.
        doc.UpdateFields();

        // Save the resulting document.
        string outputPath = Path.Combine(artifactsDir, "CustomToc.docx");
        doc.Save(outputPath);
    }
}
