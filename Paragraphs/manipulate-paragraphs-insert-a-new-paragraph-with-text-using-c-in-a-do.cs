using System;
using Aspose.Words;

class InsertParagraphExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a new paragraph with the desired text.
        // The Writeln method writes the text and then adds a paragraph break.
        builder.Writeln("This is the inserted paragraph.");

        // Save the document to a DOCX file.
        doc.Save("InsertedParagraph.docx");
    }
}
