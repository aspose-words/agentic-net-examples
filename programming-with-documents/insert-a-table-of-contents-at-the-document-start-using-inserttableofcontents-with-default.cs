using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents at the beginning of the document.
        // The switches "\\o \"1-3\" \\h \\z \\u" are the typical default settings.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Add a page break so that the TOC appears on its own page.
        builder.InsertBreak(BreakType.PageBreak);

        // Add some sample headings so the TOC has entries.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Section 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1");

        // Update fields to populate the TOC with the current headings.
        doc.UpdateFields();

        // Save the document to a file in the same folder as the executable.
        doc.Save("TableOfContents.docx");
    }
}
