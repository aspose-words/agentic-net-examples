using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write the first paragraph (no page break).
        builder.Writeln("First paragraph.");

        // Enable a forced page break before the next paragraph.
        builder.ParagraphFormat.PageBreakBefore = true;

        // Write the second paragraph; it will start on a new page.
        builder.Writeln("Second paragraph with a page break before it.");

        // Save the document to the current directory.
        doc.Save("ParagraphWithPageBreak.docx");
    }
}
