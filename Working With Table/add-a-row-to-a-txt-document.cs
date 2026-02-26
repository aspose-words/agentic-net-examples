using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class AddRowToTxtDocument
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to construct a table.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        builder.StartTable();

        // ----- First row -----
        builder.InsertCell();                     // First cell of the first row.
        builder.Write("Cell 1");                  // Add text to the cell.
        builder.InsertCell();                     // Second cell of the first row.
        builder.Write("Cell 2");
        builder.EndRow();                         // End the first row.

        // ----- New row (the row we want to add) -----
        builder.InsertCell();                     // First cell of the new row.
        builder.Write("Cell 3");
        builder.InsertCell();                     // Second cell of the new row.
        builder.Write("Cell 4");
        builder.EndRow();                         // End the new row.

        // Finish the table.
        builder.EndTable();

        // Save the document as plain text.
        TxtSaveOptions txtOptions = new TxtSaveOptions(); // Default options (paragraph break = CrLf).
        doc.Save("AddedRow.txt", txtOptions);
    }
}
