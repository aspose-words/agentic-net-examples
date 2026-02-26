using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class JoinDocumentsExample
{
    static void Main()
    {
        // Create the main (destination) document.
        Document mainDoc = new Document();

        // Use DocumentBuilder to add initial content and position the cursor.
        DocumentBuilder builder = new DocumentBuilder(mainDoc);
        builder.Writeln("This is the beginning of the main document.");

        // Load the document that will be inserted.
        Document docToInsert = new Document("Insert.docx"); // Path to the source DOCX.

        // Move the cursor to the desired insertion point.
        // In this example we insert after the existing paragraph and add a page break.
        builder.MoveToDocumentEnd();
        builder.InsertBreak(BreakType.PageBreak);

        // Insert the loaded document at the current cursor position.
        // KeepSourceFormatting preserves the original formatting of the inserted document.
        builder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        mainDoc.Save("JoinedDocument.docx");
    }
}
