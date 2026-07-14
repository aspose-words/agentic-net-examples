using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents field.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add headings that the TOC will reference.
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

        // Update all fields (including the TOC) to reflect the current document structure.
        doc.UpdateFields();

        // Rebuild the page layout so that page-number related fields in the TOC are refreshed.
        doc.UpdatePageLayout();

        // Save the document.
        doc.Save("TOC_Rebuilt.docx");
    }
}
