using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");

        // Move the builder's cursor to the start of the first paragraph.
        builder.MoveTo(doc.FirstSection.Body.FirstParagraph);

        // Insert an empty paragraph at the current cursor position.
        Paragraph emptyParagraph = builder.InsertParagraph();

        // Add more content after the inserted empty paragraph.
        builder.Writeln("Paragraph after the empty one.");

        // Save the resulting document.
        doc.Save("InsertedEmptyParagraph.docx");
    }
}
