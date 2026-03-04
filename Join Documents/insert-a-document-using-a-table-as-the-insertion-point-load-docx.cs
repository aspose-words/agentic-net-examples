using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertDocumentIntoTableCell
{
    static void Main()
    {
        // Load the destination document that contains a table.
        Document destDoc = new Document("DestinationWithTable.docx");

        // Load the source document that we want to insert.
        Document srcDoc = new Document("DocumentToInsert.docx");

        // Create a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(destDoc);

        // Move the builder's cursor to the first cell of the first table.
        // Parameters: tableIndex, rowIndex, columnIndex, cellIndex.
        // All indices are zero‑based.
        builder.MoveToCell(0, 0, 0, 0);

        // Insert the source document at the current position (inside the cell).
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the modified document.
        destDoc.Save("Result.docx");
    }
}
