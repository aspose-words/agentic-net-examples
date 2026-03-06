using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a new empty paragraph at the current cursor position.
        Paragraph newParagraph = builder.InsertParagraph();

        // Write text into the newly inserted paragraph.
        builder.Writeln("This is a newly inserted paragraph.");

        // Save the document as a DOCX file.
        doc.Save("InsertedParagraph.docx");
    }
}
