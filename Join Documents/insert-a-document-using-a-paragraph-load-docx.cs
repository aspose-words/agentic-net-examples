using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Initialize a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        // Insert an empty paragraph – this will be the insertion point.
        Paragraph insertParagraph = builder.InsertParagraph();

        // Load the source DOCX document that we want to insert.
        Document srcDoc = new Document("Source.docx");

        // Move the builder's cursor to the newly created paragraph.
        builder.MoveTo(insertParagraph);

        // Insert the source document at the cursor position, preserving its formatting.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        dstDoc.Save("Result.docx");
    }
}
