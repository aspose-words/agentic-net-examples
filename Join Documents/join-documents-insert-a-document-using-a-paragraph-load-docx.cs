using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class JoinDocumentsExample
{
    static void Main()
    {
        // Load the source document that will be inserted.
        Document srcDoc = new Document("Source.docx"); // <-- load existing DOCX

        // Create a new blank destination document.
        Document dstDoc = new Document(); // <-- create new document

        // Initialize a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        // Write some initial content and add a paragraph where the source will be inserted.
        builder.Writeln("This is the beginning of the destination document.");
        builder.InsertParagraph(); // <-- creates an empty paragraph (cursor moves to it)

        // Insert the source document at the current cursor position.
        // KeepSourceFormatting preserves the original formatting of the inserted content.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Optionally add more content after the inserted document.
        builder.Writeln();
        builder.Writeln("This is the end of the destination document.");

        // Save the combined document.
        dstDoc.Save("JoinedDocument.docx"); // <-- save to file
    }
}
