using System;
using Aspose.Words;

namespace ParagraphPageBreakExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write the first paragraph (no page break before).
            builder.Writeln("First paragraph.");

            // Enable PageBreakBefore for the next paragraph.
            builder.ParagraphFormat.PageBreakBefore = true;

            // Write the second paragraph; it will start on a new page.
            builder.Writeln("Second paragraph with a page break before it.");

            // Optionally reset the flag for subsequent paragraphs.
            builder.ParagraphFormat.PageBreakBefore = false;

            // Save the document to a file.
            doc.Save("ParagraphWithPageBreak.docx");
        }
    }
}
