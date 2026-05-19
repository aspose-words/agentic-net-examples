using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define output folder and ensure it exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents (TOC) field.
        // The switches configure the TOC to include heading levels 1‑3, add hyperlinks, hide page numbers for hidden entries, and use outline levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Populate the document with headings that the TOC will reference.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 1.1");
        builder.Writeln("Heading 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 2");
        builder.Writeln("Heading 3");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 3.1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 3.1.1");
        builder.Writeln("Heading 3.1.2");
        builder.Writeln("Heading 3.1.3");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading4;
        builder.Writeln("Heading 3.1.3.1");
        builder.Writeln("Heading 3.1.3.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 3.2");
        builder.Writeln("Heading 3.3");

        // Update all fields in the document (including the TOC field code).
        doc.UpdateFields();

        // Rebuild the page layout so that page‑related fields (PAGE, NUMPAGES, etc.) get correct values.
        doc.UpdatePageLayout();

        // Save the resulting document.
        string outputPath = Path.Combine(artifactsDir, "RebuiltToc.docx");
        doc.Save(outputPath);
    }
}
