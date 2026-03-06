using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a new paragraph containing the text "A".
        // Writeln writes the text and then adds a paragraph break.
        builder.Writeln("A");

        // Save the document to a DOCX file.
        doc.Save("Result.docx");
    }
}
