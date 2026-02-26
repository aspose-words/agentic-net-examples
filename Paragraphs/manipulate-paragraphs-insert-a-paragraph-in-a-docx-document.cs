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

        // Insert an empty paragraph at the current cursor position.
        // The method returns the inserted Paragraph node.
        Paragraph insertedParagraph = builder.InsertParagraph();

        // Write text into the newly inserted paragraph.
        // Writeln adds the text and then creates a new paragraph break.
        builder.Writeln("This is an inserted paragraph.");

        // Save the document as a DOCX file.
        doc.Save("InsertedParagraph.docx");
    }
}
