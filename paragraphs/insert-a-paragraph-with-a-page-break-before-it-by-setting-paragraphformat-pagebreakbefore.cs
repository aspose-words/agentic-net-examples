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

        // Set the PageBreakBefore flag so the next paragraph starts on a new page.
        builder.ParagraphFormat.PageBreakBefore = true;

        // Insert a paragraph with some text. The page break will be applied before this paragraph.
        builder.Writeln("This paragraph begins on a new page due to PageBreakBefore.");

        // Save the document to a file.
        doc.Save("ParagraphWithPageBreak.docx");
    }
}
