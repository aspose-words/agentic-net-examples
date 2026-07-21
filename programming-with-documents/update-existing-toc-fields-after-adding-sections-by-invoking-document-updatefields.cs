using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TOC field that captures headings level 1‑3.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add initial headings.
        InsertHeading(builder, "First Heading", StyleIdentifier.Heading1);
        InsertHeading(builder, "Subheading 1.1", StyleIdentifier.Heading2);
        InsertHeading(builder, "Subheading 1.2", StyleIdentifier.Heading2);

        // Insert a new section and more headings.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        InsertHeading(builder, "Second Section Heading", StyleIdentifier.Heading1);
        InsertHeading(builder, "Subheading 2.1", StyleIdentifier.Heading2);
        InsertHeading(builder, "Sub‑subheading 2.1.1", StyleIdentifier.Heading3);

        // Update all fields (including the TOC) to reflect the new content.
        doc.UpdateFields();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UpdatedToc.docx");
        doc.Save(outputPath);
    }

    private static void InsertHeading(DocumentBuilder builder, string text, StyleIdentifier style)
    {
        builder.ParagraphFormat.StyleIdentifier = style;
        builder.Writeln(text);
    }
}
