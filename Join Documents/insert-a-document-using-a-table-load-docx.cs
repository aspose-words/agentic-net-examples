using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertDocumentIntoTable
{
    static void Main()
    {
        // Load the source DOCX document that will be inserted.
        Document srcDoc = new Document("Source.docx");

        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Initialize a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        // Start a table in the destination document.
        builder.StartTable();

        // Insert the first cell where the source document will be placed.
        builder.InsertCell();

        // Insert the entire source document into the current cell.
        // KeepSourceFormatting preserves the original styles of the source.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // End the current row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the resulting document.
        dstDoc.Save("Result.docx");
    }
}
