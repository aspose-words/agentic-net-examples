using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define output folder and file name.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "CustomToc.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TOC that includes heading levels 1 through 3.
        // \\o "1-3" – specifies the range of heading levels.
        // \\h – makes entries clickable hyperlinks.
        // \\z – hides page numbers in web layout.
        // \\u – builds the TOC using outline levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add sample headings with levels 1, 2 and 3.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 1.1");
        builder.Writeln("Heading 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 1.1.1");
        builder.Writeln("Heading 1.1.2");

        // Update fields so the TOC reflects the added headings.
        doc.UpdateFields();

        // Save the document.
        doc.Save(outputPath);
    }
}
