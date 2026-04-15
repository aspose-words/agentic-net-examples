using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Define a folder for the generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents (TOC) field.
        // The switches configure the TOC to include heading levels 1‑3, add hyperlinks, and hide page numbers initially.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add some content with heading styles so that the TOC can pick them up.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Section 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1");
        builder.Writeln("Section 2.2");

        // First update all fields (including the TOC) – this populates the TOC entries but does not yet set correct page numbers.
        doc.UpdateFields();

        // Rebuild the page layout. This updates page‑related fields such as PAGE, NUMPAGES and also refreshes the TOC page numbers.
        doc.UpdatePageLayout();

        // Save the resulting document.
        string outFile = Path.Combine(outputDir, "RebuiltToc.docx");
        doc.Save(outFile);
    }
}
