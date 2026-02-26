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

        // Move the cursor to the end of the document.
        builder.MoveToDocumentEnd();

        // Insert an empty paragraph at the current position.
        builder.InsertParagraph();

        // Write text into the newly inserted paragraph (adds a paragraph break after the text).
        builder.Writeln("This is a newly inserted paragraph.");

        // Save the document in DOCX format.
        doc.Save("InsertedParagraph.docx");
    }
}
