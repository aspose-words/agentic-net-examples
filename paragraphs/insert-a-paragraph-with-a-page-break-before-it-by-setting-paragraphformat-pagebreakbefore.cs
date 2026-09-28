using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an initial paragraph.
        builder.Writeln("First paragraph.");

        // Insert a new empty paragraph and set a page break before it.
        builder.InsertParagraph();
        builder.CurrentParagraph.ParagraphFormat.PageBreakBefore = true;

        // Add text to the paragraph that has the page break before it.
        builder.Writeln("Paragraph with a page break before.");

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
