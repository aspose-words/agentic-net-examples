using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document (lifecycle: create)
        Document doc = new Document();

        // Initialize a DocumentBuilder to work with the document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write the first paragraph
        builder.Writeln("This is the first paragraph.");

        // Insert a new empty paragraph at the current cursor position
        // The method returns the newly created Paragraph node.
        Paragraph insertedParagraph = builder.InsertParagraph();

        // Write text into the newly inserted paragraph
        builder.Writeln("This is the inserted paragraph.");

        // Save the document to a DOCX file (lifecycle: save)
        doc.Save("InsertedParagraph.docx");
    }
}
