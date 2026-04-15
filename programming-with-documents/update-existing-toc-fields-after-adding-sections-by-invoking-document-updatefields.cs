using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents that will include headings level 1‑3.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add content to the first section.
        InsertHeading(builder, "First Section Heading 1", StyleIdentifier.Heading1);
        InsertHeading(builder, "First Section Heading 1.1", StyleIdentifier.Heading2);
        InsertHeading(builder, "First Section Heading 1.2", StyleIdentifier.Heading2);

        // Start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Add content to the second section.
        InsertHeading(builder, "Second Section Heading 1", StyleIdentifier.Heading1);
        InsertHeading(builder, "Second Section Heading 2", StyleIdentifier.Heading2);
        InsertHeading(builder, "Second Section Heading 2.1", StyleIdentifier.Heading3);

        // Update all fields in the document, including the TOC.
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save("UpdatedToc.docx");
    }

    // Helper method to write a paragraph with a specific heading style.
    private static void InsertHeading(DocumentBuilder builder, string text, StyleIdentifier style)
    {
        builder.ParagraphFormat.StyleIdentifier = style;
        builder.Writeln(text);
    }
}
