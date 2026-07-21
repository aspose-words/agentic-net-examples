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

        // Insert a Table of Contents with default switches.
        // The switches specify which heading levels to include and enable hyperlinks.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Add a page break after the TOC so headings appear on the next page.
        builder.InsertBreak(BreakType.PageBreak);

        // Add some headings that the TOC will reference.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Section 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1");

        // Update fields so the TOC reflects the headings.
        doc.UpdateFields();

        // Save the document to a file.
        doc.Save("TableOfContents.docx");
    }
}
