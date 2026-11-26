using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder for simple text insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph.");

        // Create a new paragraph.
        Paragraph paragraphWithBreak = new Paragraph(doc);
        // Set the page break before property.
        paragraphWithBreak.ParagraphFormat.PageBreakBefore = true;
        // Add some text to the paragraph.
        paragraphWithBreak.AppendChild(new Run(doc, "This paragraph starts on a new page."));

        // Append the paragraph to the document body.
        doc.FirstSection.Body.AppendChild(paragraphWithBreak);

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
