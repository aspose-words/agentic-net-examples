using System;
using Aspose.Words;

class InsertDocumentExample
{
    static void Main()
    {
        // Load the destination document (the one whose body will be the insertion point).
        Document destination = new Document("Destination.docx");

        // Load the source document that will be inserted.
        Document source = new Document("Source.docx");

        // Create a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(destination);

        // Move the builder's cursor to the first paragraph of the destination body.
        // This positions the insertion point inside the body.
        builder.MoveTo(destination.FirstSection.Body.FirstParagraph);

        // Insert the source document at the current cursor position.
        // KeepSourceFormatting preserves the original formatting of the inserted content.
        builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        destination.Save("Combined.docx");
    }
}
