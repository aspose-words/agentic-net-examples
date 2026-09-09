using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class GenerateTocExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents (TOC) field at the beginning of the document.
        // The switches configure the TOC to include heading levels 1‑3 and make entries hyperlinked.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Add a page break so that the TOC appears on its own page.
        builder.InsertBreak(BreakType.PageBreak);

        // Populate the document with headings that will be captured by the TOC.
        InsertHeading(builder, "Chapter 1 – Introduction", StyleIdentifier.Heading1);
        InsertHeading(builder, "Section 1.1 – Overview", StyleIdentifier.Heading2);
        InsertHeading(builder, "Section 1.2 – Details", StyleIdentifier.Heading2);
        InsertHeading(builder, "Chapter 2 – Usage", StyleIdentifier.Heading1);
        InsertHeading(builder, "Section 2.1 – Installation", StyleIdentifier.Heading2);
        InsertHeading(builder, "Section 2.2 – Configuration", StyleIdentifier.Heading2);
        InsertHeading(builder, "Subsection 2.2.1 – Advanced Settings", StyleIdentifier.Heading3);

        // Update all fields in the document so that the TOC reflects the headings.
        doc.UpdateFields();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the resulting document.
        string outputPath = Path.Combine(outputDir, "DocumentWithToc.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Helper method to insert a paragraph with a specific heading style.
    private static void InsertHeading(DocumentBuilder builder, string text, StyleIdentifier styleId)
    {
        builder.ParagraphFormat.StyleIdentifier = styleId;
        builder.Writeln(text);
        // Reset to normal style after inserting the heading to avoid affecting subsequent text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
    }
}
