using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a first paragraph (no page break before it).
        builder.Writeln("First paragraph.");

        // Enable a forced page break before the next paragraph.
        builder.ParagraphFormat.PageBreakBefore = true;

        // Insert the second paragraph; it will start on a new page.
        builder.Writeln("Second paragraph with a page break before it.");

        // Optional: reset the flag for any subsequent paragraphs.
        builder.ParagraphFormat.PageBreakBefore = false;

        // Save the document to the file system.
        doc.Save("ParagraphPageBreak.docx");
    }
}
