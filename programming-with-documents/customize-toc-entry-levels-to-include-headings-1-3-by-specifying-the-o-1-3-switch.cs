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

        // Insert a Table of Contents (TOC) that includes heading levels 1 through 3.
        // The switch \o "1-3" tells Word to pick up headings of levels 1‑3.
        // Additional switches:
        //   \h – make entries clickable hyperlinks,
        //   \z – hide page numbers in web layout,
        //   \u – use outline levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Insert a page break so that the headings appear after the TOC.
        builder.InsertBreak(BreakType.PageBreak);

        // Add sample headings that will be captured by the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 1.1");
        builder.Writeln("Heading 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 1.2.1");
        builder.Writeln("Heading 1.2.2");

        // Add a heading level 4 to demonstrate that it is excluded from the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading4;
        builder.Writeln("Heading 1.2.2.1 (Not in TOC)");

        // Update all fields in the document so the TOC reflects the added headings.
        doc.UpdateFields();

        // Define the output path for the generated document.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "CustomizedToc.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
