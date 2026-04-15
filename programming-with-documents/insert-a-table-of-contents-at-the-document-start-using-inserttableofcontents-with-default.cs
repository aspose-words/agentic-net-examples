using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents at the start of the document using default switches.
        // The default switches pick up headings level 1‑3, add hyperlinks, hide page numbers in web view, and use outline levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Add a page break so that the following content starts on a new page.
        builder.InsertBreak(BreakType.PageBreak);

        // Add some headings to demonstrate TOC entries.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Section 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        // Update all fields in the document so the TOC reflects the headings.
        doc.UpdateFields();

        // Save the document to the local file system.
        doc.Save("Output_Toc.docx");
    }
}
